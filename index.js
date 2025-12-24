import { characters, chat, chat_metadata, eventSource, event_types, getRequestHeaders, saveSettingsDebounced, substituteParams } from '../../../../script.js';
import { extension_settings, getContext, saveMetadataDebounced } from '../../../extensions.js';
import { SlashCommand } from '../../../slash-commands/SlashCommand.js';
import { ARGUMENT_TYPE, SlashCommandArgument, SlashCommandNamedArgument } from '../../../slash-commands/SlashCommandArgument.js';
import { SlashCommandParser } from '../../../slash-commands/SlashCommandParser.js';
import { delay, escapeRegex, isTrueBoolean } from '../../../utils.js';

const log = (...msg) => console.log('[NE]', ...msg);
/**
 * Creates a debounced function that delays invoking func until after wait milliseconds have elapsed since the last time the debounced function was invoked.
 * @param {Function} func The function to debounce.
 * @param {Number} [timeout=300] The timeout in milliseconds.
 * @returns {Function} The debounced function.
 */
export function debounceAsync(func, timeout = 300) {
    let timer;
    /**@type {Promise}*/
    let debouncePromise;
    /**@type {Function}*/
    let debounceResolver;
    return (...args) => {
        clearTimeout(timer);
        if (!debouncePromise) {
            debouncePromise = new Promise(resolve => {
                debounceResolver = resolve;
            });
        }
        timer = setTimeout(() => {
            debounceResolver(func.apply(this, args));
            debouncePromise = null;
        }, timeout);
        return debouncePromise;
    };
}

// regex cache for makeWordRegex
const regexCache = {};
// logging control: only emit verbose debug logs once per changed message
let lastLoggedMessageText = null;
let geVerboseLogging = false;

/**
 * parseBracketSpans(text)
 * - Treats single-quote ('), double-quote (") and asterisk (*) as bracket tokens.
 * - For each opening token found, scans forward to the first matching token and returns a span [start, end) where end is closeIndex+1.
 * - If no matching closer is found, returns [openIndex, text.length).
 * - Returns array of {start, end} (end exclusive).
 */
function parseBracketSpans(text) {
    const tokens = new Set(['"', '*']);
    const spans = [];
    for (let i = 0; i < text.length; i++) {
        const ch = text[i];
        if (!tokens.has(ch)) continue;
        // find next same token
        let j = text.indexOf(ch, i + 1);
        if (j === -1) {
            spans.push({ start: i, end: text.length });
            break;
        }
        spans.push({ start: i, end: j + 1 });
        i = j; // continue scanning after closer
    }
    return spans;
}

/**
 * getNonBracketSpans(text)
 * - Returns complementary spans covering text outside the bracket spans.
 * - Returns array of {start,end} (end exclusive).
 */
function getNonBracketSpans(text) {
    const brackets = parseBracketSpans(text);
    if (!brackets.length) return [{ start: 0, end: text.length }];
    const spans = [];
    let cursor = 0;
    for (const b of brackets) {
        if (b.start > cursor) {
            spans.push({ start: cursor, end: b.start });
        }
        cursor = Math.max(cursor, b.end);
    }
    if (cursor < text.length) spans.push({ start: cursor, end: text.length });
    return spans;
}

/**
 * makeWordRegex(name)
 * - Escapes name via escapeRegex and returns a cached RegExp using pattern (?:^|\\W)(escaped)(?:$|\\W) with 'gi' flags.
 */
function makeWordRegex(name) {
    const key = String(name).toLowerCase();
    if (regexCache[key]) return regexCache[key];
    const escaped = escapeRegex(name);
    const pattern = `(?:^|\\W)(${escaped})(?:$|\\W)`;
    const rx = new RegExp(pattern, 'gi');
    regexCache[key] = rx;
    return rx;
}

/**
 * countOccurrencesOutsideBrackets(name, nonBracketSpans)
 * - Counts occurrences of name (using makeWordRegex) within each non-bracket span. Returns {count, firstIndex}
 * - firstIndex is absolute index in original text of earliest match or null if none.
 */
function countOccurrencesOutsideBrackets(name, nonBracketSpans) {
    const rx = makeWordRegex(name);
    let count = 0;
    let firstIndex = null;
    for (const span of nonBracketSpans) {
        // create a fresh regex for this span to avoid shared-state issues
        const local = new RegExp(rx.source, rx.flags);
        const hay = span.text;
        if (!hay) continue;
        // use matchAll to collect matches reliably
        const matches = Array.from(hay.matchAll(local));
        if (matches.length && geVerboseLogging) {
            log('matches for', name, 'span', span.start, matches.map(m=>({ match: m[0], index: m.index })));
        }
        for (const m of matches) {
            count++;
            const absIndex = span.start + m.index;
            if (firstIndex === null || absIndex < firstIndex) firstIndex = absIndex;
        }
    }
    return { count, firstIndex };
}

/**
 * getPresentOrderedNames(lastMes, nameList)
 * - Returns ordered array of names present in lastMes according to occurrences outside bracket spans.
 * - Tie-break: descending count, then earliest unbracketed occurrence index, then master nameList index.
 * - If lastMes.is_user === true, force USER (nameList[0]) into slot 0.
 *
 * Examples:
 * // tie-break rules: higher count first, then earliest index, then master index
 * // text: "Alice says hello to Bob and Alice" -> counts: Alice=2, Bob=1 => ['Alice','Bob']
 * // text: "(Alice) Bob Alice" -> bracketed Alice ignored, counts: Alice=1 (unbracketed), Bob=1, firstIndex tie-break by earliest occurrence
 */
async function getPresentOrderedNames(lastMes, nameList) {
    const text = lastMes?.mes ?? lastMes?.message ?? lastMes?.text ?? '';
    // Determine USER name: prefer explicit custom members order (edit box) if available, else fall back to nameList[0]
    const USER_NAME = (csettings?.members && csettings.members.length) ? csettings.members[0] : nameList?.[0];
    if (geVerboseLogging) log('userNameResolution', { csettingsMembers: csettings?.members, nameListHead: nameList?.[0], USER_NAME });
    if ((!text || text.length === 0) && lastMes?.is_user) {
        return USER_NAME ? [USER_NAME] : [];
    }
    const nonBracketSpans = getNonBracketSpans(text).map(s => ({ ...s, text: text.slice(s.start, s.end) }));
    if (geVerboseLogging) log('nonBracketSpans', nonBracketSpans);
    const items = [];
    // collect counts per name for debug
    const perNameDebug = [];
    for (let i = 0; i < nameList.length; i++) {
        const name = nameList[i];
        if (csettings?.exclude?.indexOf(name.toLowerCase()) > -1) {
            perNameDebug.push({ name, count: 0, firstIndex: null, excluded: true });
            continue;
        }
        const { count, firstIndex } = countOccurrencesOutsideBrackets(name, nonBracketSpans);
        perNameDebug.push({ name, count, firstIndex, excluded: false });
        if (count > 0) items.push({ name, count, firstIndex, masterIndex: i });
    }
    // Debug log: per-name counts before sorting
    if (geVerboseLogging) log('perNameCounts', perNameDebug);
    // If message is user, ensure user exists and mark forced
    if (lastMes?.is_user) {
        if (USER_NAME) {
            const exists = items.find(it => it.name === USER_NAME);
            if (!exists) items.push({ name: USER_NAME, count: 1, firstIndex: 0, masterIndex: 0, forced: true });
            else exists.forced = true;
        }
    }
    items.sort((a, b) => {
        if (b.count !== a.count) return b.count - a.count;
        const ai = a.firstIndex == null ? Infinity : a.firstIndex;
        const bi = b.firstIndex == null ? Infinity : b.firstIndex;
        if (ai !== bi) return ai - bi;
        return a.masterIndex - b.masterIndex;
    });
    // Debug log: ordered items after sorting
    if (geVerboseLogging) log('orderedItems', items.map(it=>({ name: it.name, count: it.count, firstIndex: it.firstIndex, forced: !!it.forced })));
    // Debug: snapshot before demotion/forced adjustments
    if (geVerboseLogging) log('beforeDemotion', items.map(it=>({ name: it.name, count: it.count, firstIndex: it.firstIndex, forced: !!it.forced })));
    // Debug: demotion check state
    if (geVerboseLogging) {
        const userIdxDbg = items.findIndex(it => it.name === USER_NAME);
        log('demotionCheck', { USER_NAME, userIdx: userIdxDbg, itemsNames: items.map(it=>it.name), lastMesIsUser: !!lastMes?.is_user });
    }
    // Special rule: if not a USER message, ensure USER does not occupy slot 0 if another name exists
    if (!lastMes?.is_user && USER_NAME) {
        const userIdx = items.findIndex(it => it.name === USER_NAME);
        if (userIdx === 0 && items.length > 1) {
            // swap positions 0 and 1 so USER is at earliest position 1
            const tmp = items[0];
            items[0] = items[1];
            items[1] = tmp;
        }
    }
    // If user forced (user message), move to front
    if (lastMes?.is_user && USER_NAME) {
        const idx = items.findIndex(it => it.name === USER_NAME);
        if (idx > -1) {
            const [u] = items.splice(idx, 1);
            items.unshift(u);
        }
    }
    // Debug: snapshot after demotion/forced adjustments
    if (geVerboseLogging) log('afterDemotion', items.map(it=>({ name: it.name, count: it.count, firstIndex: it.firstIndex, forced: !!it.forced })));
    // Final priorities log
    if (geVerboseLogging) log('finalPriorities', items.map(it=>({ name: it.name, count: it.count, firstIndex: it.firstIndex, forced: !!it.forced })));
    return items.map(it => it.name);
}

// Quick sanity test (will log examples on load)
(async ()=>{
    try {
        const test1 = await getPresentOrderedNames({ mes: "Alice says hello to Bob and Alice" , is_user:false}, ['Alice','Bob','Carol']);
        log('sanity test 1 orderedNames', test1);
        const test2 = await getPresentOrderedNames({ mes: "(Alice) Bob Alice", is_user:false }, ['Alice','Bob','Carol']);
        log('sanity test 2 orderedNames', test2);
    } catch(e) {
        console.error(e);
    }
})();
/**@type {Object} */
let settings;
/**@type {Object} */
let csettings;
/**@type {String} */
let groupId;
/**@type {String} */
let chatId;
/**@type {HTMLElement} */
let root;
/**@type {HTMLElement} */
let leftArea; // DOM container covering left empty side-space
/**@type {HTMLElement} */
let rightArea; // DOM container covering right empty side-space
/**@type {HTMLElement[]} */
let imgs = [];
/**@type {String[]} */
let nameList = [];
/**@type {String[]} */
let left = [];
/**@type {String[]} */
let right = [];
/**@type {String} */
let current;
/**@type {Boolean} */
let busy = false;



/**@type {MutationObserver} */
let mo;

const updateSettingsBackground = ()=>{
    if (document.querySelector('.stge--settings .inline-drawer-content').getBoundingClientRect().height > 0 && settings.transparentMenu) {
        document.querySelector('#rm_extensions_block').style.background = 'rgba(0 0 0 / 0.5)';
    } else {
        document.querySelector('#rm_extensions_block').style.background = '';
    }
};
const initSettings = () => {
    log('initSettings');
    settings = Object.assign({
        isEnabled: true,
        numLeft: -1,
        numRight: 2,
        scaleSpeaker: 120,
        offset: 25,
        transition: 400,
        expression: 'joy',
        scaleDropoff: 3,
        transparentMenu: false,
        extensions: ['png'],
        position: 0,
        positionSingle: 100,
        placementMode: 'center', // new setting: 'center' (full-height centered) or 'width' (scale by available width)
    }, extension_settings.groupExpressions ?? {});
    extension_settings.groupExpressions = settings;

    const html = `
    <div class="stge--settings">
        <div class="inline-drawer">
            <div class="inline-drawer-toggle inline-drawer-header">
                <b>Narrator Expressions</b>
                <div class="inline-drawer-icon fa-solid fa-circle-chevron-down down"></div>
            </div>
            <div class="inline-drawer-content" style="font-size:small;">
                <div class="flex-container">
                    <label class="checkbox_label">
                        <input type="checkbox" id="stge--isEnabled" ${settings.isEnabled ? 'checked' : ''}>
                        Enable narrator expressions
                    </label>
                </div>
                <div class="flex-container">
                    <label class="checkbox_label">
                        <input type="checkbox" id="stge--transparentMenu" ${settings.transparentMenu ? 'checked' : ''}>
                        Transparent settings menu
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Position of expression images
                        <div class="stge--positionContainer">
                            Left
                            <input type="range" class="text_pole" min="0" max="100" id="stge--positionRange" value="${settings.position}">
                            Right
                            <input type="number" class="text_pole" min="0" max="100" id="stge--position" value="${settings.position}">
                            %
                        </div>
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Position of expression images with only one member
                        <div class="stge--positionContainer">
                            Left
                            <input type="range" class="text_pole" min="0" max="100" id="stge--positionSingleRange" value="${settings.positionSingle}">
                            Right
                            <input type="number" class="text_pole" min="0" max="100" id="stge--positionSingle" value="${settings.positionSingle}">
                            %
                        </div>
                    </label>
                </div>
                <!-- Placement mode controls: grouped near other position-related inputs -->
                <div class="flex-container">
                    <label>
                        Placement mode
                        <select class="text_pole" id="stge--placementMode">
                            <option value="center">Center (full height - default)</option>
                            <option value="width">Width (scale to available side width)</option>
                        </select>
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Number of characters on the left <small>(-1 = unlimited)</small>
                        <input type="number" class="text_pole" min="-1" id="stge--numLeft" value="${settings.numLeft}">
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Number of characters on the right <small>(-1 = unlimited)</small>
                        <input type="number" class="text_pole" min="-1" id="stge--numRight" value="${settings.numRight}">
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Chat path <small>(extra directory under /characters/, <strong>saved in chat</strong>)</small>
                        <input type="text" class="text_pole" id="stge--path" placeholder="Alice&Friends" value="" disabled>
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Characters to exclude <small>(comma separated list of names, <strong>saved in chat</strong>)</small>
                        <input type="text" class="text_pole" id="stge--exclude" placeholder="Alice, Bob, Carol" value="" disabled>
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Custom character list <small>(comma separated list of names, <strong>saved in chat</strong>)</small>
                        <input type="text" class="text_pole" id="stge--members" placeholder="Alice, Bob, Carol" value="" disabled>
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Scale of current speaker <small>(percentage; 100 = no change; <100 = shrink; >100 = grow)</small>
                        <input type="number" class="text_pole" min="0" id="stge--scaleSpeaker" value="${settings.scaleSpeaker}">
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Offset of characters to the side <small>(percentage; 0 = all stacked in center; <100 = overlapping; >100 = no overlap)</small>
                        <input type="number" class="text_pole" min="0" id="stge--offset" value="${settings.offset}">
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Scale dropoff <small>(percentage; 0 = no change; >0 = chars to the side get smaller; <0 = chars to the side get larger)</small>
                        <input type="number" class="text_pole" id="stge--scaleDropoff" value="${settings.scaleDropoff}">
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Animation duration <small>(milliseconds)</small>
                        <input type="number" class="text_pole" min="0" id="stge--transition" value="${settings.transition}">
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        File extensions <small>(comma-separated list, e.g. <code>png,gif,webp</code>)</small>
                        <input type="text" class="text_pole" id="stge--extensions" value="${settings.extensions.join(',')}">
                    </label>
                </div>
                <div class="flex-container">
                    <label>
                        Default expression to be used
                        <select class="text_pole" id="stge--expression"></select>
                    </label>
                </div>
            </div>
        </div>
    </div>
`;
    $('#extensions_settings').append(html);
    window.addEventListener('click', ()=>{
        updateSettingsBackground();
    });
    document.querySelector('#stge--isEnabled').addEventListener('click', ()=>{
        settings.isEnabled = document.querySelector('#stge--isEnabled').checked;
        saveSettingsDebounced();
        restart();
    });
    document.querySelector('#stge--transparentMenu').addEventListener('click', ()=>{
        settings.transparentMenu = document.querySelector('#stge--transparentMenu').checked;
        saveSettingsDebounced();
    });
    document.querySelector('#stge--positionRange').addEventListener('input', ()=>{
        settings.position = document.querySelector('#stge--positionRange').value;
        document.querySelector('#stge--position').value = settings.position;
        saveSettingsDebounced();
        if (namesCount > 1 && root) {
            root.style.setProperty('--position', settings.position);
        }
    });
    document.querySelector('#stge--position').addEventListener('input', ()=>{
        settings.position = document.querySelector('#stge--position').value;
        document.querySelector('#stge--positionRange').value = settings.position;
        saveSettingsDebounced();
        if (namesCount > 1 && root) {
            root.style.setProperty('--position', settings.position);
        }
    });
    document.querySelector('#stge--positionSingleRange').addEventListener('input', ()=>{
        settings.positionSingle = document.querySelector('#stge--positionSingleRange').value;
        document.querySelector('#stge--positionSingle').value = settings.positionSingle;
        saveSettingsDebounced();
        if (namesCount == 1 && root) {
            root.style.setProperty('--position', settings.positionSingle);
        }
    });
    document.querySelector('#stge--positionSingle').addEventListener('input', ()=>{
        settings.positionSingle = document.querySelector('#stge--positionSingle').value;
        document.querySelector('#stge--positionSingleRange').value = settings.positionSingle;
        saveSettingsDebounced();
        if (namesCount == 1 && root) {
            root.style.setProperty('--position', settings.positionSingle);
        }
    });

    // initialize placement mode selector value
    const placementSel = document.querySelector('#stge--placementMode');
    if (placementSel) placementSel.value = settings.placementMode;
    placementSel?.addEventListener('change', ()=>{
        // save setting, debounce, and immediately update UI via CSS variables without restarting
        settings.placementMode = document.querySelector('#stge--placementMode').value;
        saveSettingsDebounced();
        if (root) {
            root.style.setProperty('--placement-mode', settings.placementMode);
            // update side sizes/areas immediately so image wrappers move to correct containers
            updateSideSizes();
        }
    });

    document.querySelector('#stge--numLeft').addEventListener('input', ()=>{
        settings.numLeft = Number(document.querySelector('#stge--numLeft').value);
        saveSettingsDebounced();
    });
    document.querySelector('#stge--numRight').addEventListener('input', ()=>{
        settings.numRight = Number(document.querySelector('#stge--numRight').value);
        saveSettingsDebounced();
    });
    document.querySelector('#stge--path').addEventListener('input', ()=>{
        csettings.path = document.querySelector('#stge--path').value;
        chat_metadata.groupExpressions = csettings;
        saveMetadataDebounced();
    });
    document.querySelector('#stge--exclude').addEventListener('input', ()=>{
        csettings.exclude = document.querySelector('#stge--exclude').value.toLowerCase().split(/\s*,\s*/).filter(it=>it.length);
        chat_metadata.groupExpressions = csettings;
        saveMetadataDebounced();
    });
    document.querySelector('#stge--members').addEventListener('input', ()=>{
        csettings.members = document.querySelector('#stge--members').value.split(/\s*,\s*/).filter(it=>it.length);
        chat_metadata.groupExpressions = csettings;
        saveMetadataDebounced();
    });
    document.querySelector('#stge--scaleSpeaker').addEventListener('input', ()=>{
        settings.scaleSpeaker = Number(document.querySelector('#stge--scaleSpeaker').value);
        saveSettingsDebounced();
        root?.style.setProperty('--scale-speaker', String(settings.scaleSpeaker));
    });
    document.querySelector('#stge--offset').addEventListener('input', ()=>{
        settings.offset = Number(document.querySelector('#stge--offset').value);
        saveSettingsDebounced();
        root?.style.setProperty('--offset', String(settings.offset));
    });
    document.querySelector('#stge--scaleDropoff').addEventListener('input', ()=>{
        settings.scaleDropoff = Number(document.querySelector('#stge--scaleDropoff').value);
        saveSettingsDebounced();
        root?.style.setProperty('--scale-dropoff', String(settings.scaleDropoff));
    });
    document.querySelector('#stge--transition').addEventListener('input', ()=>{
        settings.transition = Number(document.querySelector('#stge--transition').value);
        saveSettingsDebounced();
        root?.style.setProperty('--transition', String(settings.transition));
    });
    document.querySelector('#stge--extensions').addEventListener('input', ()=>{
        settings.extensions = document.querySelector('#stge--extensions').value?.split(/,\s*/);
        saveSettingsDebounced();
        restart();
    });
    const sel = document.querySelector('#stge--expression');
    const exp = [
        'admiration',
        'amusement',
        'anger',
        'annoyance',
        'approval',
        'caring',
        'confusion',
        'curiosity',
        'desire',
        'disappointment',
        'disapproval',
        'disgust',
        'embarrassment',
        'excitement',
        'fear',
        'gratitude',
        'grief',
        'joy',
        'love',
        'nervousness',
        'neutral',
        'optimism',
        'pride',
        'realization',
        'relief',
        'remorse',
        'sadness',
        'surprise',
    ];
    exp.forEach(e=>{
        const opt = document.createElement('option'); {
            opt.value = e;
            opt.textContent = e;
            opt.selected = (settings.expression ?? 'joy') == e;
            sel.append(opt);
        }
    });
    sel.addEventListener('change', ()=>{
        settings.expression = sel.value;
        saveSettingsDebounced();
    });
};

/**
 * updateSideSizes()
 * - Finds the chatbox element (#sheld) and measures available empty horizontal space to the left/right of it.
 * - Updates CSS variables on the root element for --left-space, --right-space and --placement-mode so CSS can position/scale images.
 * - Also updates the width of the leftArea/rightArea DOM containers so wrappers appended into them are clipped/positioned correctly.
 * - This function is intentionally idempotent and cheap so it can be called on window.resize and before rendering updates.
 */
function updateSideSizes() {
    if (!root) return;
    // Find the main chat container. Prefer getElementById but fallback to querySelector as requested.
    const sheld = document.getElementById('sheld') || document.querySelector('#sheld');
    let leftSpace = 0;
    let rightSpace = 0;
    if (sheld) {
        const rect = sheld.getBoundingClientRect();
        leftSpace = Math.max(0, Math.floor(rect.left));
        rightSpace = Math.max(0, Math.floor(window.innerWidth - rect.right));
    } else {
        // No chatbox found - treat full viewport as available (split half/half) to avoid zero-width areas
        leftSpace = Math.floor(window.innerWidth / 2);
        rightSpace = window.innerWidth - leftSpace;
    }
    // expose pixel values as CSS variables for use by CSS rules
    root.style.setProperty('--left-space', `${leftSpace}px`);
    root.style.setProperty('--right-space', `${rightSpace}px`);
    root.style.setProperty('--placement-mode', settings.placementMode);
    // update DOM area widths so appended wrappers are inside the correct side area
    if (leftArea) leftArea.style.width = `${leftSpace}px`;
    if (rightArea) rightArea.style.width = `${rightSpace}px`;
    if (geVerboseLogging) log('updateSideSizes', { leftSpace, rightSpace });
}

const chatChanged = async ()=>{
    log('chatChanged');
    namesCount = -1;
    const context = getContext();

    csettings = Object.assign({
        exclude: [],
        members: [],
        emotes: {},
    }, chat_metadata.groupExpressions ?? {});
    chat_metadata.groupExpressions = csettings;
    log(chat_metadata);
    document.querySelector('#stge--path').disabled = context.chatId == null;
    document.querySelector('#stge--path').value = csettings.path ?? '';
    document.querySelector('#stge--exclude').disabled = context.chatId == null;
    document.querySelector('#stge--exclude').value = csettings.exclude?.join(', ') ?? '';
    document.querySelector('#stge--members').disabled = context.chatId == null;
    document.querySelector('#stge--members').value = csettings.members?.join(', ') ?? '';

    if (true || context.groupId) {
        await restart();
    } else {
        end();
    }
};

const groupUpdated = (...args) => {
    log('GROUP UPDATED', args);
    namesCount = -1;
};

const messageRendered = async () => {
    log('messageRendered');
    while (settings.isEnabled && (groupId || true)) {
        if (!busy) {
            updateSettingsBackground();
            // Ensure side area sizes are recalculated before layout is applied so wrappers are appended into correct containers
            if (root) updateSideSizes();
            await updateMembers();
            const lastMes = chat.toReversed().find(it=>!it.is_system);
            const lastCharMes = chat.toReversed().find(it=>!it.is_user && !it.is_system && nameList.find(o=>it.name == o));
            // Decide whether to emit verbose debug logs for this message (only once per changed message)
            const messageTextForLog = lastMes?.mes ?? lastMes?.message ?? lastMes?.text ?? '';
            if (messageTextForLog !== lastLoggedMessageText) {
                geVerboseLogging = true;
                lastLoggedMessageText = messageTextForLog;
            } else {
                geVerboseLogging = false;
            }

            // New presence & ordering logic (narrator/DM mode): compute ordered names based on unbracketed occurrences
            const orderedNames = await getPresentOrderedNames(lastMes, nameList);
            if (geVerboseLogging) log('orderedNames', orderedNames);
            const slots = orderedNames.slice(0, 4);
            // expose how many images are visible so CSS can adapt layouts for 1/2/3 images
            if (root) root.setAttribute('data-visible-count', String(Math.max(0, Math.min(4, slots.length))));

            // debug: print slots and wrapper state when verbose logging is enabled
            if (geVerboseLogging) {
                try {
                    log('messageRendered slots:', slots);
                    log('wrappers:', imgs.map(w=>({ name: w.getAttribute('data-character'), attached: !!w.closest('.stge--root'), parent: w.parentElement?.className })));
                    log('side widths (px):', { left: leftArea?.getBoundingClientRect?.().width, right: rightArea?.getBoundingClientRect?.().width });
                } catch(e) { log('debug log error', e); }
            }

            // Clean previous "last" markers
            imgs.filter(it=>it.classList.contains('stge--last')).forEach(it=>it.classList.remove('stge--last'));

            // Show/hide and assign slot ordering
            for (const wrapper of imgs) {
                const name = wrapper.getAttribute('data-character');
                const slotIndex = slots.indexOf(name);
                if (slotIndex >= 0) {
                    // assign corner slot via --order
                    wrapper.style.setProperty('--order', String(slotIndex));
                    // enter animation if not in root
                    if (!wrapper.closest('.stge--root')) {
                        wrapper.classList.add('stge--exit');
                        // Map wrappers to side areas according to visibleCount so original layout is preserved:
                        // 1 -> slot0 left full-height
                        // 2 -> slot0 left, slot1 right
                        // 3 -> slot0 left full-height, slot1 top-right, slot2 bottom-right
                        // 4 -> slot0 top-left, slot1 bottom-left, slot2 top-right, slot3 bottom-right
                        if (geVerboseLogging) log('placing wrapper', name, { slotIndex, slots });
                        const visibleCount = slots.length;
                        let targetArea = null;
                        if (visibleCount === 1) {
                            targetArea = leftArea;
                        } else if (visibleCount === 2) {
                            targetArea = (slotIndex === 0) ? leftArea : rightArea;
                        } else if (visibleCount === 3) {
                            // slotIndex 0 -> left, 1/2 -> right
                            targetArea = (slotIndex === 0) ? leftArea : rightArea;
                        } else {
                            // visibleCount >=4: distribute 0/1 -> left, 2/3 -> right
                            targetArea = (slotIndex <= 1) ? leftArea : rightArea;
                        }
                        // append into appropriate side-area container so wrapper occupies the empty side-space rather than full viewport
                        targetArea?.append(wrapper);
                        await delay(50);
                        wrapper.classList.remove('stge--exit');
                    }
                    wrapper.classList.remove('stge--hidden');
                } else {
                    // hide it
                    wrapper.style.removeProperty('--order');
                    // if currently attached to root (or one of its side areas), animate exit and remove from DOM (element object remains in imgs)
                    if (wrapper.closest('.stge--root')) {
                        wrapper.classList.add('stge--exit');
                        await delay(settings.transition + 150);
                        wrapper.remove();
                    } else {
                        // keep in memory but mark hidden
                        wrapper.classList.add('stge--hidden');
                    }
                }
            }

            // Mark the last response visually: if last message is not from user, and primary speaker corresponds to first slot
            if (lastMes?.is_user === false) {
                const primary = slots[0];
                if (primary && lastCharMes && primary === lastCharMes.name) {
                    const wrap = imgs.find(it=>it.getAttribute('data-character') == primary && it.closest('.stge--root'));
                    if (wrap) wrap.classList.add('stge--last');
                }
            }
        }
        await delay(Math.max(settings.transition + 100, 1000));
    }
};
// eventSource.on(event_types.CHARACTER_MESSAGE_RENDERED, ()=>messageRendered());




const findImage = async(name, expression = null) => {
    for (const ext of settings.extensions) {
        const url = csettings.exclude ? `/characters/${csettings.path}/${name}/${expression ?? settings.expression}.${ext}` : `/characters/${name}/${expression ?? settings.expression}.${ext}`;
        const resp = await fetch(url, {
            method: 'HEAD',
            headers: getRequestHeaders(),
        });
        if (resp.ok) {
            return url;
        }
    }
    if (expression && expression != settings.expression) {
        return await findImage(name);
    }
};
let namesCount = -1;
const updateMembers = async()=>{
    if (busy) return;
    busy = true;
    try {
        const context = getContext();
        const names = [];
        if (csettings.members?.length) {
            const members = getOrderFromText(csettings.members);
            names.push(...members);
            names.push(...csettings.members.filter(m=>!names.find(it=>it == m)));
        } else if (groupId) {
            const group = context.groups.find(it=>it.id == groupId);
            const members = group.members.map(m=>characters.find(c=>c.avatar == m)).filter(it=>it);
            names.push(...getOrder(members.map(it=>it.name)).filter(it=>csettings.exclude?.indexOf(it.toLowerCase()) == -1));
            names.push(...members.filter(m=>!names.find(it=>it == m.name)).map(it=>it.name).filter(it=>csettings.exclude?.indexOf(it.toLowerCase()) == -1));
        } else if (context.characterId) {
            names.push(characters[context.characterId].name);
        }
        if (namesCount != names.length) {
            namesCount = names.length;
            if (names.length == 1) {
                root?.style.setProperty('--position', settings.positionSingle ?? '100');
            } else {
                root?.style.setProperty('--position', settings.position);
            }
        }
        const removed = nameList.filter(it=>names.indexOf(it) == -1);
        const added = names.filter(it=>nameList.indexOf(it) == -1);
        for (const name of removed) {
            nameList.splice(nameList.indexOf(name), 1);
            let idx = imgs.findIndex(it=>it.getAttribute('data-character') == name);
            const img = imgs.splice(idx, 1)[0];
            idx = left.indexOf(name);
            if (idx > -1) {
                left.splice(idx, 1);
            } else {
                idx = right.indexOf(name);
                if (idx > -1) {
                    right.splice(idx, 1);
                } else {
                    current = null;
                }
            }
            img.classList.add('stge--exit');
            await delay(settings.transition + 150);
            img.remove();
        }
        const purgatory = [];
        while (settings.numLeft != -1 && left.length > settings.numLeft) {
            purgatory.push(left.pop());
        }
        while (settings.numRight != -1 && right.length > settings.numRight) {
            purgatory.push(right.pop());
        }
        for (const name of added) {
            nameList.push(name);
            if (!current) {
                current = name;
            } else if ((left.length < settings.numLeft || settings.numLeft == -1) && (left.length <= right.length || right.length >= settings.numRight)) {
                left.push(name);
            } else if (right.length < settings.numRight || settings.numRight == -1) {
                right.push(name);
            }
            const wrap = document.createElement('div'); {
                imgs.push(wrap);
                wrap.classList.add('stge--wrapper');
                wrap.setAttribute('data-character', name);
                const img = document.createElement('img'); {
                    img.classList.add('stge--img');
                    let tc = chat_metadata.triggerCards ?? {};
                    if (!tc?.isEnabled) {
                        tc = {};
                    }
                    img.src = await findImage(tc?.costumes?.[name] ?? name, csettings[name]?.emote);
                    wrap.append(img);
                    if (geVerboseLogging) log('created wrapper for', name, 'src', img.src);
                }
            }
        }
        for (const name of purgatory) {
            if (!current) {
                current = name;
            } else if ((left.length < settings.numLeft || settings.numLeft == -1) && (left.length <= right.length || right.length >= settings.numRight)) {
                left.push(name);
            } else if (right.length < settings.numRight || settings.numRight == -1) {
                right.push(name);
            } else {
                const wrap = imgs.find(it=>it.getAttribute('data-character') == name);
                if (wrap) {
                    wrap.classList.add('stge--exit');
                    await delay(settings.transition + 150);
                    wrap.remove();
                }
            }
        }
        const queue = nameList.filter(it=>left.indexOf(it)==-1 && right.indexOf(it) == -1 && it != current);
        while (queue.length > 0 && (settings.numLeft == -1 || settings.numRight == -1 || left.length < settings.numLeft || right.length < settings.numRight || !current)) {
            const name = queue.pop();
            if (!current) {
                current = name;
            } else if ((left.length < settings.numLeft || settings.numLeft == -1) && (left.length <= right.length || right.length >= settings.numRight)) {
                left.push(name);
            } else if (right.length < settings.numRight || settings.numRight == -1) {
                right.push(name);
            }
        }
    } catch (ex) {
        console.error('[NE]', ex);
    }
    busy = false;
};
eventSource.on(event_types.CHAT_CHANGED, ()=>(chatChanged(),null));
eventSource.on(event_types.GROUP_UPDATED, (...args)=>groupUpdated(...args));
// eventSource.on(event_types.USER_MESSAGE_RENDERED, ()=>messageRendered());




const getOrder = (members)=>{
    const o = [];

    const mesList = chat.filter(it=>!it.is_system && !it.is_user && members.indexOf(it.name) > -1).toReversed();
    for (const mes of mesList) {
        if (o.indexOf(mes.name) == -1) {
            o.unshift(mes.name);
            if (o.length >= members.length) {
                break;
            }
        }
    }
    return o;
};
const getOrderFromText = (members)=>{
    members = [...members];
    const o = [];
    const regex = members.map(it=>[it, new RegExp(`(?:^|\\W)(${escapeRegex(it)})(?:$|\\W)`)]).reduce((dict,cur)=>(dict[cur[0]] = cur[1], dict), {});
    const mesList = chat.filter(it=>!it.is_system && !it.is_user).toReversed();
    for (const mes of mesList) {
        const mesmem = [];
        for (const m of members) {
            if (regex[m].test(mes.mes)) {
                const match = regex[m].exec(mes.mes);
                mesmem.push([m, match]);
            }
        }
        mesmem.sort((a,b)=>a[1].index - b[1].index);
        o.push(...mesmem.map(it=>it[0]).filter((it,idx,list)=>idx == list.indexOf(it)));
        for (const m of mesmem) {
            members.splice(members.indexOf(m[0]), 1);
        }
    }
    return o;
};
let restarting = false;
const restart = debounceAsync(async()=>{
    if (restarting) return;
    restarting = true;
    log('restart');
    end();
    await delay(Math.max(550, settings.transition + 150));
    await start();
    restarting = false;
});
const start = async()=>{
    if (!settings.isEnabled) return;
    log('start');
    document.querySelector('#expression-wrapper').style.opacity = '0';
    root = document.createElement('div'); {
        root.classList.add('stge--root');
        root.style.setProperty('--scale-speaker', settings.scaleSpeaker);
        root.style.setProperty('--offset', settings.offset);
        root.style.setProperty('--transition', settings.transition);
        root.style.setProperty('--scale-dropoff', settings.scaleDropoff);
        root.style.setProperty('--position', namesCount == 1 ? (settings.positionSingle ?? '100') : settings.position);
        root.style.setProperty('--placement-mode', settings.placementMode);
        document.body.append(root);
    }
    // Create left/right side-area containers that cover the empty side spaces of the viewport.
    // These containers receive wrapper elements so they occupy only the empty space beside the centered chatbox.
    leftArea = document.createElement('div');
    // give container both the generic area class and the left-specific class so CSS rules apply
    leftArea.classList.add('stge--area', 'stge--left-area');
    leftArea.style.position = 'absolute';
    leftArea.style.left = '0';
    leftArea.style.top = '0';
    leftArea.style.bottom = '0';
    // allow wrappers/images to overflow when needed; CSS expects visible
    leftArea.style.overflow = 'visible';
    root.append(leftArea);

    rightArea = document.createElement('div');
    // give container both the generic area class and the right-specific class so CSS rules apply
    rightArea.classList.add('stge--area', 'stge--right-area');
    rightArea.style.position = 'absolute';
    rightArea.style.right = '0';
    rightArea.style.top = '0';
    rightArea.style.bottom = '0';
    rightArea.style.overflow = 'visible';
    root.append(rightArea);

    // Listen for resize to keep side sizes in sync with the centered chatbox
    window.addEventListener('resize', updateSideSizes);
    // initial measurement
    updateSideSizes();

    const context = getContext();
    groupId = context.groupId;
    chatId = context.chatid;
    messageRendered();
    mo.observe(document.querySelector('#expression-wrapper'), { childList:true, subtree:true, attributes:true });
    document.querySelector('#expression-wrapper').style.opacity = '0';
};
const end = ()=>{
    log('end');
    mo.disconnect();
    groupId = null;
    chatId = null;
    current = null;
    nameList = [];
    left = [];
    right = [];
    // remove resize listener and cleanup side-area elements
    window.removeEventListener('resize', updateSideSizes);
    leftArea?.remove();
    rightArea?.remove();
    leftArea = null;
    rightArea = null;
    root?.remove();
    root = null;
    while (imgs.length > 0) {
        imgs.pop();
    }
    document.querySelector('#expression-wrapper').style.opacity = '';
    log('/end');
};

const init = ()=>{
    log('init');
    initSettings();
    mo = new MutationObserver(async(muts)=>{
        if (busy) return;
        const lastCharMes = chat.toReversed().find(it=>!it.is_user && !it.is_system && nameList.find(o=>it.name == o));
        const img = imgs.find(it=>it.getAttribute('data-character') == lastCharMes?.name);
        if (img && document.querySelector('#expression-image').src) {
            const src = document.querySelector('#expression-image').src;
            const parts = src.split('/');
            const name = parts.slice(parts.indexOf('characters') + 1, -1).map(it=>decodeURIComponent(it));
            if (csettings.emotes[name[0]]?.isLocked) return;
            const img = imgs.find(it=>it.getAttribute('data-character') == name[0])?.querySelector('.stge--img');
            if (img) {
                let tc = chat_metadata.triggerCards ?? {};
                if (!tc?.isEnabled) {
                    tc = {};
                }
                img.src = await findImage(tc?.costumes?.[name] ?? name.join('/'), parts.at(-1).replace(/^(.+)\.[^.]+$/, '$1'));
            }
        }
    });
};
init();



SlashCommandParser.addCommandObject(SlashCommand.fromProps({ name: 'ge-members',
    /**
     *
     * @param {*} args
     * @param {string} value
     */
    callback: (args, value)=>{
        const inp = /**@type {HTMLInputElement}*/(document.querySelector('#stge--members'));
        try {
            inp.value = JSON.parse(value).join(', ');
            inp.dispatchEvent(new Event('input'));
        } catch {
            // no (or no valid) value provided, nothing to do
        }
        return JSON.stringify(csettings.members ?? []);
    },
    unnamedArgumentList: [
        SlashCommandArgument.fromProps({ description: 'list of names, [] to clear',
            typeList: ARGUMENT_TYPE.LIST,
        }),
    ],
    returns: 'current custom member list',
    helpString: `
        <div>
            Update the custom member list for Narrator Expressions.
        </div>
        <div>
            Leave the unnamed argument blank to just return the current custom member list.
        </div>
        <div>
            <strong>Examples:</strong>
            <ul>
                <li><pre><code class="language-stscript">/ge-members ["Alice", "Bob"]</code></pre></li>
                <li><pre><code class="language-stscript">/ge-members | /echo</code></pre></li>
            </ul>
        </div>
    `,
}));

SlashCommandParser.addCommandObject(SlashCommand.fromProps({ name: 'ge-emote',
    /**
     * @param {{name:string, lock:string, clear:string}} args
     * @param {string} value
     */
    callback: async(args, value)=>{
        const name = args.name ?? substituteParams('{{char}}');
        if (isTrueBoolean(args.clear)) {
            delete csettings.emotes[name];
            saveMetadataDebounced();
            return '';
        }
        if (value?.length) {
            csettings.emotes[name] = {
                emote: value,
                isLocked: isTrueBoolean(args.lock ?? 'false'),
            };
            saveMetadataDebounced();
            const img = imgs.find(it=>it.getAttribute('data-character') == name)?.querySelector('.stge--img');
            if (img) {
                let tc = chat_metadata.triggerCards ?? {};
                if (!tc?.isEnabled) {
                    tc = {};
                }
                img.src = await findImage(tc?.costumes?.[name] ?? name, value);
            }
        }
        const result = csettings.emotes[name];
        if (result) return JSON.stringify(result);
        return '';
    },
    namedArgumentList: [
        SlashCommandNamedArgument.fromProps({ name: 'name',
            description: 'name of the member',
            defaultValue: '{{char}}',
        }),
        SlashCommandNamedArgument.fromProps({ name: 'lock',
            description: 'true: emote cannot be automatically changed by expressions extension',
            typeList: [ARGUMENT_TYPE.BOOLEAN],
            defaultValue: 'false',
        }),
        SlashCommandNamedArgument.fromProps({ name: 'clear',
            description: 'true: remove the manually set emote',
            typeList: [ARGUMENT_TYPE.BOOLEAN],
            defaultValue: 'false',
        }),
    ],
    unnamedArgumentList: [
        SlashCommandArgument.fromProps({ description: 'expression',
        }),
    ],
    returns: 'current emote',
    helpString: `
        <div>
            Set the emote / expression for a member.
        </div>
        <div>
            Leave the unnamed argument blank to just return the current expression.
        </div>
        <div>
            <strong>Examples:</strong>
            <ul>
                <li><pre><code class="language-stscript">/ge-emote joy</code></pre></li>
                <li><pre><code class="language-stscript">/ge-emote name=Alice lock=true joy</code></pre></li>
                <li><pre><code class="language-stscript">/ge-emote name=Bob clear=true</code></pre></li>
            </ul>
        </div>
    `,
}));
