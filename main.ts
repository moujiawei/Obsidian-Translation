import { App, Editor, Plugin, PluginSettingTab, Setting, Notice, Modal, TFile } from "obsidian";
// @ts-ignore
import * as CryptoJS from "crypto-js";
const { exec } = require('child_process'); // <-- 移到这里
import * as fs from 'fs';
import * as path from 'path';
import * as XLSX from 'xlsx';
import { markdownTable } from 'markdown-table';
const OpenCC = require('opencc-js');
const t2s = OpenCC.Converter({ from: 'tw', to: 'cn' });


    
    

// === 全局唯一词汇本弹窗实例 ===
let globalVocabBookModal: VocabBookModal | null = null;

interface YoudaoSettings {
    textAppKey: string;
    textAppSecret: string;
    imageAppKey: string;
    imageAppSecret: string;
    textTargetLang: string;
    imageTargetLang: string;
    textTranslationColor: string;
    imageTranslationColor: string;
    textTranslationMode: string;
    imageTranslationMode: string;
    serverPath: string; // 服务器路径 (server.exe)
    serverPort: string; // 端口
    serverPathMode?: 'local' | 'cloud'; // 新增
    dictWordLimitEn?: number; // 英文分界
    dictWordLimitZh?: number; // 中文分界
    isBiDirection: boolean;
    sleepInterval?: number; // 新增：翻译间隔ms
    microsoftKey?: string; // 新增：微软密钥
    microsoftRegion?: string; // 新增：微软位置/区域
    microsoftEndpoint?: string; // 新增：微软终结点
    translateModel?: string; // 新增：翻译模型选择
    microsoftSleepInterval?: number; // 微软多行翻译间隔
    microsoftImageKey?: string; // 新增：图片翻译微软密钥
    microsoftImageRegion?: string; // 新增：图片翻译微软位置/区域
    microsoftImageEndpoint?: string; // 新增：图片翻译微软终结点
    dictServicePath?: string; // 新增：dict_service 路径

}

const DEFAULT_SETTINGS: YoudaoSettings = {
    textAppKey: "",
    textAppSecret: "",
    imageAppKey: "",
    imageAppSecret: "",
    textTargetLang: "en",
    imageTargetLang: "en",
    textTranslationColor: "#1a73e8",
    imageTranslationColor: "#e67e22",
    textTranslationMode: "merge",
    imageTranslationMode: "merge",
    serverPath: "",
    serverPort: "4000",
    serverPathMode: 'local', // 默认本地磁盘
    dictWordLimitEn: 3, // 英文分界默认3
    dictWordLimitZh: 4,  // 中文分界默认4
    isBiDirection: false,
    sleepInterval: 250,
    microsoftKey: '',
    microsoftRegion: '',
    microsoftEndpoint: '',
    translateModel: 'youdao',
    microsoftSleepInterval: 250,
    microsoftImageKey: '',
    microsoftImageRegion: '',
    microsoftImageEndpoint: '',

};


function getFindPidCmd(port: number) {
    if (process.platform === 'win32') {
        return `netstat -ano | findstr :${port}`;
    } else {
        return `lsof -i :${port} | grep LISTEN`;
    }
}
function getKillCmd(pid: string) {
    if (process.platform === 'win32') {
        return `taskkill /PID ${pid} /F`;
    } else {
        return `kill -9 ${pid}`;
    }
}

function getDictCmd(args: string) {
    if (process.platform === 'win32') {
        return `api_dict.exe ${args}`;
    } else if (process.platform === 'darwin') {
        return `./api_dict_mac ${args}`; // 你需要准备Mac版
    } else {
        return `./api_dict_linux ${args}`; // 你需要准备Linux版
    }
}


function saveServerPathHistory(plugin: YoudaoTranslatePlugin, newPath: string) {
    if (!newPath) return;
    let history = plugin.serverHistory.serverPathHistory || [];
    history = [newPath, ...history.filter(p => p !== newPath)];
    if (history.length > 10) history = history.slice(0, 10);
    plugin.serverHistory.serverPathHistory = history;
    plugin.saveServerHistory();
}

// 全局缓存快捷键
let selectionHotkeyCombo: any = null;
let lineHotkeyCombo: any = null;
function cacheYoudaoHotkeys(app: App) {
    console.log('[YoudaoPlugin] cacheYoudaoHotkeys called');
    function getCombo(commandId: string) {
        const hotkeyManager = (app as any).hotkeyManager;
        console.log('[YoudaoPlugin] getCombo hotkeyManager:', hotkeyManager);
        if (!hotkeyManager) {
            console.log('[YoudaoPlugin] hotkeyManager 不可用');
            return null;
        }
        const custom = hotkeyManager.customKeys?.[commandId] || [];
        const builtIn = hotkeyManager.hotkeys?.[commandId] || [];
        console.log(`[YoudaoPlugin] getCombo commandId=${commandId} custom:`, custom, 'builtIn:', builtIn);
        const all = [...custom, ...builtIn];
        console.log(`[YoudaoPlugin] getCombo commandId=${commandId} all:`, all);
        if (all.length === 0) {
            console.log(`[YoudaoPlugin] 命令 ${commandId} 没有快捷键`);
            return null;
        }
        const h = all[0];
        console.log(`[YoudaoPlugin] getCombo commandId=${commandId} h:`, h);
        const combo = {
            ctrl: (h.modifiers || []).includes('Ctrl'),
            shift: (h.modifiers || []).includes('Shift'),
            alt: (h.modifiers || []).includes('Alt'),
            key: h.key?.toLowerCase()
        };
        console.log(`[YoudaoPlugin] 命令 ${commandId} 快捷键:`, combo);
        return combo;
    }
    selectionHotkeyCombo = getCombo('youdao-translate-selection-to-english');
    lineHotkeyCombo = getCombo('youdao-translate-current-line');
    console.log('[YoudaoPlugin] 缓存快捷键 selectionHotkeyCombo:', selectionHotkeyCombo, 'lineHotkeyCombo:', lineHotkeyCombo);
}

// 类型和默认值提前声明
interface VocabBookData {
    vocabBook: any[];
}
interface VocabTrashData {
    vocabBookTrash: any[];
}
interface ServerHistoryData {
    serverPathHistory: string[];
    dictServicePathHistory?: string[];
}
const DEFAULT_VOCAB_DATA: VocabBookData = { vocabBook: [] };
const DEFAULT_TRASH_DATA: VocabTrashData = { vocabBookTrash: [] };
const DEFAULT_SERVER_HISTORY: ServerHistoryData = { serverPathHistory: [] };

// 工具函数：确保 data 目录存在
async function ensureDataDir(app: App) {
    const dir = '.obsidian/plugins/Obsidian Translation/data';
    console.log('[YoudaoPlugin] 检查数据目录:', dir);
    if (!(await app.vault.adapter.exists(dir))) {
        console.log('[YoudaoPlugin] 创建数据目录:', dir);
        await app.vault.adapter.mkdir(dir);
    } else {
        console.log('[YoudaoPlugin] 数据目录已存在:', dir);
    }
}

// 1. 定义统一的翻译服务接口（扩展词典查词）
interface TranslateService {
    translateText(text: string, from: string, to: string): Promise<string | null>;
    lookupWord?(word: string, port: string): Promise<any>;
    translateImage?(params: {
        app: App,
        file: TFile,
        imageAppKey: string,
        imageAppSecret: string,
        textAppKey: string,
        textAppSecret: string,
        to: string,
        color: string,
        mode: string,
        port: string,
        plugin?: any
    }): Promise<void>;
}

// 2. 有道适配器实现
class YoudaoAdapter implements TranslateService {
    appKey: string;
    appSecret: string;
    app: App;
    port: string;
    constructor(appKey: string, appSecret: string, app: App, port: string) {
        this.appKey = appKey;
        this.appSecret = appSecret;
        this.app = app;
        this.port = port;
    }
    async translateText(text: string, from: string, to: string): Promise<string | null> {
        return await youdaoTranslate(text, this.appKey, this.appSecret, to, this.app, this.port);
    }
    async lookupWord(word: string, port: string): Promise<any> {
        return await youdaoDictLookup(word, port);
    }
    async translateImage(params: {
        app: App,
        file: TFile,
        imageAppKey: string,
        imageAppSecret: string,
        textAppKey: string,
        textAppSecret: string,
        to: string,
        color: string,
        mode: string,
        port: string,
        plugin?: any
    }): Promise<void> {
        return await translateImageFile(
            params.app,
            params.file,
            params.imageAppKey,
            params.imageAppSecret,
            params.textAppKey,
            params.textAppSecret,
            params.to,
            params.color,
            params.mode,
            params.port,
            params.plugin
        );
    }
}

// 新增：自动查找可用端口（递增）
async function findAvailablePort(startPort: number, maxTries = 20): Promise<number> {
    let port = startPort;
    for (let i = 0; i < maxTries; i++) {
        const inUse = await checkPortInUse(port);
        if (!inUse) return port;
        port++;
    }
    throw new Error('未找到可用端口');
}


// 新增：微软词典查词和例句API
async function microsoftDictLookup(text: string, from: string, to: string, msAdapter: MicrosoftAdapter): Promise<any> {
    if (!msAdapter.key || !msAdapter.region || !msAdapter.endpoint) return null;
    const url = msAdapter.endpoint.replace(/\/$/, '') + `/dictionary/lookup?api-version=3.0&from=${from}&to=${to}`;
    const resp = await fetch(url, {
        method: 'POST',
        headers: {
            'Ocp-Apim-Subscription-Key': msAdapter.key,
            'Ocp-Apim-Subscription-Region': msAdapter.region,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify([{ Text: text }])
    });
    return await resp.json();
}
async function microsoftDictExamples(text: string, translation: string, from: string, to: string, msAdapter: MicrosoftAdapter): Promise<any> {
    if (!msAdapter.key || !msAdapter.region || !msAdapter.endpoint) return null;
    const url = msAdapter.endpoint.replace(/\/$/, '') + `/dictionary/examples?api-version=3.0&from=${from}&to=${to}`;
    const resp = await fetch(url, {
        method: 'POST',
        headers: {
            'Ocp-Apim-Subscription-Key': msAdapter.key,
            'Ocp-Apim-Subscription-Region': msAdapter.region,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify([{ Text: text, Translation: translation }])
    });
    return await resp.json();
}
async function waitPortRelease(port: number, timeout = 5000, interval = 100) {
    const start = Date.now();
    while (await checkPortInUse(port)) {
        if (Date.now() - start > timeout) return false;
        await new Promise(res => setTimeout(res, interval));
    }
    return true;
}


// 2. 表格预览辅助函数
function renderHtmlTable(headers: string[], rows: any[][], highlightRows: number[], highlightCols: number[], onRowClick: (row: number) => void, onColClick: (col: number) => void, enableRowClick: boolean = true, startIndex: number = 0): HTMLElement {
    const table = document.createElement('table');
    table.style.width = '100%';
    table.style.borderCollapse = 'collapse';
    table.style.marginTop = '16px';
    table.style.fontSize = '14px';
    table.style.background = '#fff';
    table.style.border = '1px solid #ccc';
    table.style.tableLayout = 'fixed'; // 关键：固定布局
    // 列宽数组，初始均为120px
    const defaultColWidth = 120;
    const colWidths: number[] = Array(headers.length).fill(defaultColWidth);
    // 行高数组，初始均为auto
    const rowHeights: number[] = Array(rows.length).fill(undefined);
    // --- 表头 ---
    const thead = document.createElement('thead');
    const tr = document.createElement('tr');
    headers.forEach((h, colIdx) => {
        const th = document.createElement('th');
        th.textContent = h;
        th.style.border = '1px solid #ccc';
        th.style.padding = '4px 8px';
        th.style.background = '#f7f7f7';
        th.style.position = 'relative';
        th.style.width = colWidths[colIdx] + 'px'; // 初始化宽度
        // 拖拽手柄
        const resizer = document.createElement('div');
        resizer.style.position = 'absolute';
        resizer.style.right = '0';
        resizer.style.top = '0';
        resizer.style.width = '6px';
        resizer.style.height = '100%';
        resizer.style.cursor = 'col-resize';
        resizer.style.userSelect = 'none';
        resizer.style.zIndex = '10';
        resizer.style.background = 'transparent';
        resizer.onmouseenter = () => { resizer.style.background = '#1a73e8'; };
        resizer.onmouseleave = () => { resizer.style.background = 'transparent'; };
        resizer.onmousedown = (e) => {
            e.preventDefault();
            const startX = e.clientX;
            const startWidth = th.offsetWidth;
            function onMouseMove(ev: MouseEvent) {
                const dx = ev.clientX - startX;
                let newWidth = Math.max(40, startWidth + dx);
                th.style.width = newWidth + 'px';
                colWidths[colIdx] = newWidth;
                // 同步所有td宽度
                Array.from(table.querySelectorAll('tbody tr')).forEach(rowTr => {
                    const cell = rowTr.children[colIdx] as HTMLElement;
                    if (cell) cell.style.width = newWidth + 'px';
                });
            }
            function onMouseUp() {
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
            }
            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        };
        th.appendChild(resizer);
        tr.appendChild(th);
    });
    thead.appendChild(tr);
    table.appendChild(thead);
    // --- 表体 ---
    const tbody = document.createElement('tbody');
    rows.forEach((row, rowIdx) => {
        const realRowNum = startIndex + rowIdx + 1;
        const tr = document.createElement('tr');
        if (rowHeights[rowIdx]) tr.style.height = rowHeights[rowIdx] + 'px';
        if (highlightRows.includes(realRowNum)) tr.classList.add('vocab-row-highlight');
        if (enableRowClick) {
            tr.onclick = (e) => {
                onRowClick(realRowNum);
            };
        }
        row.forEach((cell, colIdx) => {
            const td = document.createElement('td');
            td.textContent = String(cell ?? '');
            td.style.border = '1px solid #ccc';
            td.style.padding = '4px 8px';
            td.style.width = colWidths[colIdx] + 'px'; // 初始化宽度
            if (highlightCols.includes(colIdx + 1)) td.classList.add('vocab-row-highlight');
            td.onclick = (e) => {
                if (!enableRowClick) e.stopPropagation(); // 只在列模式阻止冒泡
                onColClick(colIdx + 1);
            };
            tr.appendChild(td);
        });
        // 行高拖拽手柄
        const rowResizer = document.createElement('div');
        rowResizer.style.position = 'absolute';
        rowResizer.style.left = '0';
        rowResizer.style.right = '0';
        rowResizer.style.bottom = '0';
        rowResizer.style.height = '6px';
        rowResizer.style.cursor = 'row-resize';
        rowResizer.style.userSelect = 'none';
        rowResizer.style.zIndex = '10';
        rowResizer.style.background = 'transparent';
        tr.style.position = 'relative';
        rowResizer.onmouseenter = () => { rowResizer.style.background = '#1a73e8'; };
        rowResizer.onmouseleave = () => { rowResizer.style.background = 'transparent'; };
        rowResizer.onmousedown = (e) => {
            e.preventDefault();
            const startY = e.clientY;
            const startHeight = tr.offsetHeight;
            function onMouseMove(ev: MouseEvent) {
                const dy = ev.clientY - startY;
                let newHeight = Math.max(24, startHeight + dy);
                tr.style.height = newHeight + 'px';
                rowHeights[rowIdx] = newHeight;
            }
            function onMouseUp() {
                document.removeEventListener('mousemove', onMouseMove);
                document.removeEventListener('mouseup', onMouseUp);
            }
            document.addEventListener('mousemove', onMouseMove);
            document.addEventListener('mouseup', onMouseUp);
        };
        tr.appendChild(rowResizer);
        tbody.appendChild(tr);
    });
    table.appendChild(tbody);
    return table;
}

// 3. csv/tsv/md/xlsx/ods 解析辅助函数（这里只实现 csv/tsv/md，xlsx/ods 可后续补充）
function parseCsvTsv(text: string, delimiter: string = ','): { headers: string[], rows: string[][] } {
    const lines = text.split(/\r?\n/).filter(Boolean);
    const headers = lines[0].split(delimiter);
    const rows = lines.slice(1).map(l => l.split(delimiter));
    return { headers, rows };
}
function parseMarkdownTable(text: string): { headers: string[], rows: string[][] } {
    // 只保留以|开头的行
    const lines = text.split(/\r?\n/).filter(l => l.trim().startsWith('|'));
    if (lines.length < 2) return { headers: [], rows: [] };
    // 表头，允许空格
    let headers = lines[0].split('|').map(s => s.trim());
    // 允许表头首尾空格
    if (headers[0] === '') headers = headers.slice(1);
    if (headers[headers.length - 1] === '') headers = headers.slice(0, -1);
    // 分隔线
    const sepIdx = 1;
    // 数据行
    const rows: string[][] = [];
    for (let i = 2; i < lines.length; i++) {
        let row = lines[i].split('|').map(s => s.trim());
        if (row[0] === '') row = row.slice(1);
        if (row[row.length - 1] === '') row = row.slice(0, -1);
        // 补齐到和表头一样多
        while (row.length < headers.length) row.push('');
        while (row.length > headers.length) row = row.slice(0, headers.length);
        rows.push(row);
    }
    // 处理全空行（即所有单元格都是空字符串）
    // 不过滤任何行
    // 处理全空列：不做过滤，渲染时自然显示
    return { headers, rows };
}


export default class YoudaoTranslatePlugin extends Plugin {
    settings: YoudaoSettings;
    vocabData: VocabBookData;
    trashData: VocabTrashData;
    serverHistory: ServerHistoryData;
    private serverProcess: any = null;
    _refreshStatus?: () => void;
    private _currentServerPort?: number;
    private _currentServerPath?: string;

    async onload() {
        console.log('[YoudaoPlugin] 插件 onload 启动');
        await this.loadAllData();
        this.addSettingTab(new YoudaoSettingTab(this.app, this));

        // 全局交互拦截：有最小化弹窗时，任何交互都 notice 并阻止
        let lastNoticeTime = 0;
        function globalBlocker(e: Event) {
            if (hasMinimizedModal()) {
                const now = Date.now();
                if (now - lastNoticeTime > 1000) {
                    new Notice("请先还原/关闭顶层弹窗，再进行其他操作");
                    lastNoticeTime = now;
                }
                if (e.cancelable) e.preventDefault();
                e.stopImmediatePropagation && e.stopImmediatePropagation();
                e.stopPropagation();
                return false;
            }
        }
        // 监听主交互事件
        ["mousedown", "wheel", "keydown", "touchstart"].forEach(type => {
            window.addEventListener(type, globalBlocker, true); // 捕获阶段
        });

        this.addCommand({
            id: "youdao-translate-selection-to-english",
            name: "翻译选中内容（弹窗显示）",
            editorCallback: async (editor: Editor) => {
                if (hasMinimizedModal()) {
                    new Notice("请先还原/关闭顶层弹窗，再进行其他操作");
                    return;
                }
                const selectedText = (window.getSelection()?.toString() || '') || editor.getSelection();
                if (!selectedText) {
                    new Notice("请先选中要翻译的内容");
                    return;
                }
                new Notice("正在翻译...");
                if (this.settings.translateModel === 'composite') {
                    // === 新增：多行直接机器翻译 ===
                    if (selectedText.split(/\r?\n/).length > 1) {
                        const lines = selectedText.split(/\r?\n/).filter(line => line.trim().length > 0);
                        await unifiedBatchTranslateForMicrosoft({
                            app: this.app,
                            items: lines,
                            getTargetLang: this.settings.isBiDirection
                                ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans'
                                : () => this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang,
                            msAdapter: new MicrosoftAdapter(
                                this.settings.microsoftKey || '',
                                this.settings.microsoftRegion || '',
                                this.settings.microsoftEndpoint || '',
                                this.app
                            ),
                            color: this.settings.textTranslationColor,
                            mode: this.settings.textTranslationMode,
                            sleepInterval: this.settings.microsoftSleepInterval ?? 250,
                            plugin: this
                        });
                        return;
                    }
                    // === 分界逻辑 ===
                    const mainLang = detectMainLangSmart(selectedText);
                    const isZh = mainLang === 'zh';
                    let useDict = false;
                    if (isZh) {
                        const zhCount = (selectedText.match(/[\u4e00-\u9fa5]/g) || []).length;
                        const zhLimit = this.settings.dictWordLimitZh ?? 4;
                        useDict = zhCount > 0 && zhCount <= zhLimit;
                    } else {
                        const wordCount = selectedText.trim().split(/\s+/).length;
                        const enLimit = this.settings.dictWordLimitEn ?? 3;
                        useDict = wordCount > 0 && wordCount <= enLimit;
                    }
                    if (useDict) {
                        // 复合模式下，查词用本地 api_dict 接口，UI风格与有道一致
                        try {
                            const resp = await fetch(`http://127.0.0.1:${this.settings.serverPort}/api/dict?q=${encodeURIComponent(selectedText)}`);
                            if (!resp.ok) throw new Error('本地词典API请求失败');
                            const dictData = await resp.json();
                            console.log('[YoudaoPlugin] 本地词典返回数据:', dictData); // <--- 加这一行
                            // 判断是否有基本释义
                            let hasBasic = false;
                            if (Array.isArray(dictData.definitions) && dictData.definitions.length > 0) {
                                hasBasic = true;
                            }
                            let machineTranslation = '';
                            if (!hasBasic) {
                                // 机器翻译目标语言逻辑（与超分界一致）
                                let tLang = 'en';
                                if (this.settings.isBiDirection) {
                                    tLang = isZh ? 'en' : 'zh-Hans';
                                } else {
                                    tLang = this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang;
                                }
                                const from = isZh ? 'zh-Hans' : 'en';
                                machineTranslation = await new MicrosoftAdapter(
                                    this.settings.microsoftKey || '',
                                    this.settings.microsoftRegion || '',
                                    this.settings.microsoftEndpoint || '',
                                    this.app
                                ).translateText(selectedText, from, tLang) || '';
                            }
                            // 适配有道UI三大板块
                            const html = renderDictResult({
                                definitions: dictData.definitions || [],
                                examples: dictData.examples || [],
                                phrases: dictData.phrases || []
                            }, isZh ? 'zh' : 'en', this.settings.textTranslationColor, this.settings.serverPort, machineTranslation, selectedText);
                            new DictResultModal(this.app, html, this).open();
                            return;
                        } catch (e) {
                            new Notice('本地词典释义获取失败：' + e.message);
                            return;
                        }
                    }
                    // 超过分界线，走微软机器翻译
                    // 多行处理
                    const lines = selectedText.split(/\r?\n/).filter(line => line.trim().length > 0);
                    if (lines.length > 1 || (lines.length === 1 && /[。！？.!?]/.test(lines[0]))) {
                        await unifiedBatchTranslateForMicrosoft({
                            app: this.app,
                            items: lines,
                            getTargetLang: this.settings.isBiDirection
                                ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans'
                                : () => this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang,
                            msAdapter: new MicrosoftAdapter(
                                this.settings.microsoftKey || '',
                                this.settings.microsoftRegion || '',
                                this.settings.microsoftEndpoint || '',
                                this.app
                            ),
                            color: this.settings.textTranslationColor,
                            mode: this.settings.textTranslationMode,
                            sleepInterval: this.settings.microsoftSleepInterval ?? 250,
                            plugin: this
                        });
                        return;
                    }
                    // 单行机器翻译
                    let targetLang = 'en';
                    if (this.settings.isBiDirection) {
                        targetLang = isZh ? 'en' : 'zh-Hans';
                    } else {
                        targetLang = this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang;
                    }
                    if ((targetLang === 'zh-Hans' && isZh) || (targetLang === 'en' && !isZh)) {
                        new TranslateResultModal(this.app, selectedText, selectedText, this.settings.textTranslationColor, this.settings.textTranslationMode, this).open();
                        return;
                    }
                    const from = isZh ? 'zh-Hans' : 'en';
                    const translated = await new MicrosoftAdapter(
                        this.settings.microsoftKey || '',
                        this.settings.microsoftRegion || '',
                        this.settings.microsoftEndpoint || '',
                        this.app
                    ).translateText(selectedText, from, targetLang);
                    console.log('[翻译日志][MicrosoftAdapter] 原文:', selectedText, 'from:', from, 'to:', targetLang, '返回:', translated);
                    if (translated) {
                        new TranslateResultModal(this.app, selectedText, translated, this.settings.textTranslationColor, this.settings.textTranslationMode, this).open();
                    } else {
                        new Notice('微软翻译失败');
                    }
                    return;
                }
                await handleDictOrTranslate(selectedText, this.app, this.settings, this);
            }
        });

        this.addCommand({
            id: "youdao-translate-selected-image",
            name: "翻译选中图片（弹窗显示）",
            editorCallback: async (editor: Editor) => {
                if (hasMinimizedModal()) {
                    new Notice("请先还原/关闭顶层弹窗，再进行其他操作");
                    return;
                }
                let selectedText = editor.getSelection();
                if (!selectedText) {
                    selectedText = window.getSelection()?.toString() || '';
                }
                if (!selectedText) {
                    new Notice("请先选中图片链接（![[xxx.png]]）");
                    return;
                }
                const match = selectedText.match(/\[\[(.+?\.(png|jpg|jpeg|gif|bmp))/i);
                if (!match) {
                    new Notice("选中的内容不是图片链接");
                    return;
                }
                const imageName = match[1];
                const files = this.app.vault.getFiles();
                const file = files.find(f => f.name === imageName);
                if (!file) {
                    new Notice(`未找到图片文件: ${imageName}`);
                    return;
                }
                // 接口化调用
                if (this.settings.translateModel === 'composite') {
                    // 复合翻译模式下，走微软图片翻译
                    const msAdapter = new MicrosoftAdapter(
                        this.settings.microsoftKey || '',
                        this.settings.microsoftRegion || '',
                        this.settings.microsoftEndpoint || '',
                        this.app
                    );
                    await microsoftTranslateImageFile(
                        this.app,
                        file,
                        msAdapter,
                        this.settings.imageTranslationColor,
                        this.settings.imageTranslationMode,
                        this
                    );
                } else {
                    // 有道翻译模式
                    const translator: TranslateService = new YoudaoAdapter(
                        this.settings.textAppKey,
                        this.settings.textAppSecret,
                        this.app,
                        this.settings.serverPort
                    );
                    await translator.translateImage?.({
                        app: this.app,
                        file,
                        imageAppKey: this.settings.imageAppKey,
                        imageAppSecret: this.settings.imageAppSecret,
                        textAppKey: this.settings.textAppKey,
                        textAppSecret: this.settings.textAppSecret,
                        to: this.settings.isBiDirection ? '' : this.settings.imageTargetLang,
                        color: this.settings.imageTranslationColor,
                        mode: this.settings.imageTranslationMode,
                        port: this.settings.serverPort,
                        plugin: this
                    });
                }
            }
        });

        this.addCommand({
            id: "youdao-translate-current-line",
            name: "翻译当前行（弹窗显示）",
            editorCallback: async (editor: Editor) => {
                if (hasMinimizedModal()) {
                    new Notice("请先还原/关闭顶层弹窗，再进行其他操作");
                    return;
                }
                const cursor = editor.getCursor();
                const lineText = editor.getLine(cursor.line);
                if (!lineText || !(lineText ?? '').toString().trim()) {
                    new Notice("当前行为空");
                    return;
                }
                new Notice("正在翻译...");
                if (this.settings.translateModel === 'composite') {
                    // === 新增：多行直接机器翻译 ===
                    if (lineText.split(/\r?\n/).length > 1) {
                        const lines = lineText.split(/\r?\n/).filter(line => line.trim().length > 0);
                        await unifiedBatchTranslateForMicrosoft({
                            app: this.app,
                            items: lines,
                            getTargetLang: this.settings.isBiDirection
                                ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans'
                                : () => this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang,
                            msAdapter: new MicrosoftAdapter(
                                this.settings.microsoftKey || '',
                                this.settings.microsoftRegion || '',
                                this.settings.microsoftEndpoint || '',
                                this.app
                            ),
                            color: this.settings.textTranslationColor,
                            mode: this.settings.textTranslationMode,
                            sleepInterval: this.settings.microsoftSleepInterval ?? 250,
                            plugin: this
                        });
                        return;
                    }
                    // === 分界逻辑 ===
                    const mainLang = detectMainLangSmart(lineText);
                    const isZh = mainLang === 'zh';
                    let useDict = false;
                    if (isZh) {
                        const zhCount = (lineText.match(/[\u4e00-\u9fa5]/g) || []).length;
                        const zhLimit = this.settings.dictWordLimitZh ?? 4;
                        useDict = zhCount > 0 && zhCount <= zhLimit;
                    } else {
                        const wordCount = lineText.trim().split(/\s+/).length;
                        const enLimit = this.settings.dictWordLimitEn ?? 3;
                        useDict = wordCount > 0 && wordCount <= enLimit;
                    }
                    if (useDict) {
                        // 复合模式下，查词用本地 api_dict 接口，UI风格与有道一致
                        try {
                            const resp = await fetch(`http://127.0.0.1:${this.settings.serverPort}/api/dict?q=${encodeURIComponent(lineText)}`);
                            if (!resp.ok) throw new Error('本地词典API请求失败');
                            const dictData = await resp.json();
                            console.log('[YoudaoPlugin] 本地词典返回数据:', dictData); // <--- 加这一行
                            // 判断是否有基本释义
                            let hasBasic = false;
                            if (Array.isArray(dictData.definitions) && dictData.definitions.length > 0) {
                                hasBasic = true;
                            }
                            let machineTranslation = '';
                            if (!hasBasic) {
                                // 机器翻译目标语言逻辑（与超分界一致）
                                let tLang = 'en';
                                if (this.settings.isBiDirection) {
                                    tLang = isZh ? 'en' : 'zh-Hans';
                                } else {
                                    tLang = this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang;
                                }
                                const from = isZh ? 'zh-Hans' : 'en';
                                machineTranslation = await new MicrosoftAdapter(
                                    this.settings.microsoftKey || '',
                                    this.settings.microsoftRegion || '',
                                    this.settings.microsoftEndpoint || '',
                                    this.app
                                ).translateText(lineText, from, tLang) || '';
                            }
                            // 适配有道UI三大板块
                            const html = renderDictResult({
                                definitions: dictData.definitions || [],
                                examples: dictData.examples || [],
                                phrases: dictData.phrases || []
                            }, isZh ? 'zh' : 'en', this.settings.textTranslationColor, this.settings.serverPort, machineTranslation, lineText);
                            new DictResultModal(this.app, html, this).open();
                            return;
                        } catch (e) {
                            new Notice('本地词典释义获取失败：' + e.message);
                            return;
                        }

                    }
                    // 超过分界线，走微软机器翻译
                    // 多行处理
                    const lines = lineText.split(/\r?\n/).filter(line => line.trim().length > 0);
                    if (lines.length > 1 || (lines.length === 1 && /[。！？.!?]/.test(lines[0]))) {
                        await unifiedBatchTranslateForMicrosoft({
                            app: this.app,
                            items: lines,
                            getTargetLang: this.settings.isBiDirection
                                ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans'
                                : () => this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang,
                            msAdapter: new MicrosoftAdapter(
                                this.settings.microsoftKey || '',
                                this.settings.microsoftRegion || '',
                                this.settings.microsoftEndpoint || '',
                                this.app
                            ),
                            color: this.settings.textTranslationColor,
                            mode: this.settings.textTranslationMode,
                            sleepInterval: this.settings.microsoftSleepInterval ?? 250,
                            plugin: this
                        });
                        return;
                    }
                    // 单行机器翻译
                    let targetLang = 'en';
                    if (this.settings.isBiDirection) {
                        targetLang = isZh ? 'en' : 'zh-Hans';
                    } else {
                        targetLang = this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang;
                    }
                    if ((targetLang === 'zh-Hans' && isZh) || (targetLang === 'en' && !isZh)) {
                        new TranslateResultModal(this.app, lineText, lineText, this.settings.textTranslationColor, this.settings.textTranslationMode, this).open();
                        return;
                    }
                    const from = isZh ? 'zh-Hans' : 'en';
                    const translated = await new MicrosoftAdapter(
                        this.settings.microsoftKey || '',
                        this.settings.microsoftRegion || '',
                        this.settings.microsoftEndpoint || '',
                        this.app
                    ).translateText(lineText, from, targetLang);
                    console.log('[翻译日志][MicrosoftAdapter] 原文:', lineText, 'from:', from, 'to:', targetLang, '返回:', translated);
                    if (translated) {
                        new TranslateResultModal(this.app, lineText, translated, this.settings.textTranslationColor, this.settings.textTranslationMode, this).open();
                    } else {
                        new Notice('微软翻译失败');
                    }
                    return;
                }
                await handleDictOrTranslate(lineText, this.app, this.settings, this);
            }
        });

        this.addCommand({
            id: "youdao-translate-file-title",
            name: "翻译当前笔记标题（弹窗显示）",
            callback: async () => {
                if (hasMinimizedModal()) {
                    new Notice("请先还原/关闭顶层弹窗，再进行其他操作");
                    return;
                }
                const file = this.app.workspace.getActiveFile();
                if (!file) {
                    new Notice("未找到当前笔记");
                    return;
                }
                const title = file.basename;
                if (this.settings.translateModel === 'composite') {
                    // === 新增：多行直接机器翻译 ===
                    if (title.split(/\r?\n/).length > 1) {
                        const lines = title.split(/\r?\n/).filter(line => line.trim().length > 0);
                        await unifiedBatchTranslateForMicrosoft({
                            app: this.app,
                            items: lines,
                            getTargetLang: this.settings.isBiDirection
                                ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans'
                                : () => this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang,
                            msAdapter: new MicrosoftAdapter(
                                this.settings.microsoftKey || '',
                                this.settings.microsoftRegion || '',
                                this.settings.microsoftEndpoint || '',
                                this.app
                            ),
                            color: this.settings.textTranslationColor,
                            mode: this.settings.textTranslationMode,
                            sleepInterval: this.settings.microsoftSleepInterval ?? 250,
                            plugin: this
                        });
                        return;
                    }
                    // === 分界逻辑 ===
                    const mainLang = detectMainLangSmart(title);
                    const isZh = mainLang === 'zh';
                    let useDict = false;
                    if (isZh) {
                        const zhCount = (title.match(/[\u4e00-\u9fa5]/g) || []).length;
                        const zhLimit = this.settings.dictWordLimitZh ?? 4;
                        useDict = zhCount > 0 && zhCount <= zhLimit;
                    } else {
                        const wordCount = title.trim().split(/\s+/).length;
                        const enLimit = this.settings.dictWordLimitEn ?? 3;
                        useDict = wordCount > 0 && wordCount <= enLimit;
                    }
                    if (useDict) {
                        // 复合模式下，查词用本地 api_dict 接口，UI风格与有道一致
                        try {
                            const resp = await fetch(`http://127.0.0.1:${this.settings.serverPort}/api/dict?q=${encodeURIComponent(title)}`);
                            if (!resp.ok) throw new Error('本地词典API请求失败');
                            const dictData = await resp.json();
                            console.log('[YoudaoPlugin] 本地词典返回数据:', dictData); // <--- 加这一行
                            // 判断是否有基本释义
                            let hasBasic = false;
                            if (Array.isArray(dictData.definitions) && dictData.definitions.length > 0) {
                                hasBasic = true;
                            }
                            let machineTranslation = '';
                            if (!hasBasic) {
                                // 机器翻译目标语言逻辑（与超分界一致）
                                let tLang = 'en';
                                if (this.settings.isBiDirection) {
                                    tLang = isZh ? 'en' : 'zh-Hans';
                                } else {
                                    tLang = this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang;
                                }
                                const from = isZh ? 'zh-Hans' : 'en';
                                machineTranslation = await new MicrosoftAdapter(
                                    this.settings.microsoftKey || '',
                                    this.settings.microsoftRegion || '',
                                    this.settings.microsoftEndpoint || '',
                                    this.app
                                ).translateText(title, from, tLang) || '';
                            }
                            // 适配有道UI三大板块
                            const html = renderDictResult({
                                definitions: dictData.definitions || [],
                                examples: dictData.examples || [],
                                phrases: dictData.phrases || []
                            }, isZh ? 'zh' : 'en', this.settings.textTranslationColor, this.settings.serverPort, machineTranslation, title);
                            new DictResultModal(this.app, html, this).open();
                            return;
                        } catch (e) {
                            new Notice('本地词典释义获取失败：' + e.message);
                            return;
                        }
                    }
                    // 超过分界线，走微软机器翻译
                    // 多行处理
                    const lines = title.split(/\r?\n/).filter(line => line.trim().length > 0);
                    if (lines.length > 1 || (lines.length === 1 && /[。！？.!?]/.test(lines[0]))) {
                        await unifiedBatchTranslateForMicrosoft({
                            app: this.app,
                            items: lines,
                            getTargetLang: this.settings.isBiDirection
                                ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans'
                                : () => this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang,
                            msAdapter: new MicrosoftAdapter(
                                this.settings.microsoftKey || '',
                                this.settings.microsoftRegion || '',
                                this.settings.microsoftEndpoint || '',
                                this.app
                            ),
                            color: this.settings.textTranslationColor,
                            mode: this.settings.textTranslationMode,
                            sleepInterval: this.settings.microsoftSleepInterval ?? 250,
                            plugin: this
                        });
                        return;
                    }
                    // 单行机器翻译
                    let targetLang = 'en';
                    if (this.settings.isBiDirection) {
                        targetLang = isZh ? 'en' : 'zh-Hans';
                    } else {
                        targetLang = this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang;
                    }
                    if ((targetLang === 'zh-Hans' && isZh) || (targetLang === 'en' && !isZh)) {
                        new TranslateResultModal(this.app, title, title, this.settings.textTranslationColor, this.settings.textTranslationMode, this).open();
                        return;
                    }
                    const from = isZh ? 'zh-Hans' : 'en';
                    const translated = await new MicrosoftAdapter(
                        this.settings.microsoftKey || '',
                        this.settings.microsoftRegion || '',
                        this.settings.microsoftEndpoint || '',
                        this.app
                    ).translateText(title, from, targetLang);
                    console.log('[翻译日志][MicrosoftAdapter] 原文:', title, 'from:', from, 'to:', targetLang, '返回:', translated);
                    if (translated) {
                        new TranslateResultModal(this.app, title, translated, this.settings.textTranslationColor, this.settings.textTranslationMode, this).open();
                    } else {
                        new Notice('微软翻译失败');
                    }
                    return;
                }
                await handleDictOrTranslate(title, this.app, this.settings, this);
            }
        });

        // 全局禁止 modal-bg 拦截鼠标事件
        const style = document.createElement('style');
        style.textContent = `
.modal-bg { pointer-events: none !important; }
.modal { z-index: 99999 !important; pointer-events: auto !important; }
`;
        document.head.appendChild(style);

        // 在插件 onload 里注册命令（或复用已有命令），增加自动加入词汇本逻辑
        this.addCommand({
            id: "youdao-dict-add-to-vocab",
            name: "词典释义并加入词汇本",
            hotkeys: [], // 用户可自定义
            editorCallback: async (editor: Editor) => {
                const selectedText = (window.getSelection()?.toString() || '') || editor.getSelection();
                if (!selectedText) {
                    new Notice("请先选中要加入词汇本的内容");
                    return;
                }
                // 只要包含中文就禁止收录
                if (/[\u4e00-\u9fa5]/.test(selectedText)) {
                    new Notice("被选中目标不能包含中文");
                    return;
                }
                // 判断英文单词数是否超限
                const wordCount = selectedText.trim().split(/\s+/).length;
                const enLimit = this.settings.dictWordLimitEn ?? 3;
                if (wordCount > enLimit) {
                    new Notice(`单词/短语数量超过限制（当前${wordCount}，最大${enLimit}）`);
                    return;
                }
                let dictData;
                let word = selectedText.trim();
                let translation = '';
                let example = '';
                let notes = '';
                // === 新增：根据模式切换查词逻辑 ===
                if (this.settings.translateModel === 'composite') {
                    // 复合模式：本地 api_dict 查词
                    try {
                        const resp = await fetch(`http://127.0.0.1:${this.settings.serverPort}/api/dict?q=${encodeURIComponent(word)}`);
                        if (!resp.ok) throw new Error('本地词典API请求失败');
                        dictData = await resp.json();
                    } catch (e) {
                        new Notice('本地词典释义获取失败：' + e.message);
                        return;
                    }
                    // 解析 definitions
                    let definitions = Array.isArray(dictData.definitions) ? dictData.definitions : [];
                    if (definitions.length > 0) {
                        translation = definitions.map((item: any) => item.en || item.zh || item.sense || '').filter(Boolean).join('\n');
                    }
                    // 解析例句
                    let examples = Array.isArray(dictData.examples) ? dictData.examples : [];
                    if (examples.length > 0) {
                        example = examples.map((ex: any) => `${ex.en || ''}\n${ex.zh || ''}`.trim()).filter(Boolean).join('\n\n');
                    }
                    // 解析短语/网络释义
                    let phrases = Array.isArray(dictData.phrases) ? dictData.phrases : [];
                    if (phrases.length > 0) {
                        notes = phrases.map((item: any) => {
                            if (typeof item === 'string') return item;
                            return `${item.key || ''}: ${item.trans || ''}`;
                        }).join('\n');
                    }
                } else {
                    // 有道模式：原有逻辑
                    try {
                        const translator: TranslateService = new YoudaoAdapter(
                            this.settings.textAppKey,
                            this.settings.textAppSecret,
                            this.app,
                            this.settings.serverPort
                        );
                        dictData = await translator.lookupWord?.(word, this.settings.serverPort);
                    } catch (e) {
                        new Notice('词典释义获取失败：' + e.message);
                        return;
                    }
                    let w = null;
                    if (dictData.ec && dictData.ec.word && dictData.ec.word.length > 0) {
                        w = dictData.ec.word[0];
                    } else if (dictData.ce && dictData.ce.word && dictData.ce.word.length > 0) {
                        w = dictData.ce.word[0];
                    } else if (dictData.ee && dictData.ee.word && dictData.ee.word.length > 0) {
                        w = dictData.ee.word[0];
                    }
                    let trsList = [];
                    if (w && w.trs && w.trs.length > 0) {
                        trsList = w.trs;
                    } else if (dictData.ec && dictData.ec.word && dictData.ec.word[0] && dictData.ec.word[0].trs && dictData.ec.word[0].trs.length > 0) {
                        trsList = dictData.ec.word[0].trs;
                    } else if (dictData.ce && dictData.ce.word && dictData.ce.word[0] && dictData.ce.word[0].trs && dictData.ce.word[0].trs.length > 0) {
                        trsList = dictData.ce.word[0].trs;
                    } else if (dictData.ee && dictData.ee.word && dictData.ee.word[0] && dictData.ee.word[0].trs && dictData.ee.word[0].trs.length > 0) {
                        trsList = dictData.ee.word[0].trs;
                    }
                    if (trsList.length > 0) {
                        const lines: string[] = [];
                        trsList.forEach((tr: any) => {
                            if (tr.tr && tr.tr[0] && tr.tr[0].l && tr.tr[0].l.i) {
                                const item = tr.tr[0].l.i;
                                function extractStrings(obj: any): string[] {
                                    let result: string[] = [];
                                    if (typeof obj === 'string') {
                                        result.push(obj);
                                    } else if (Array.isArray(obj)) {
                                        obj.forEach(sub => {
                                            result = result.concat(extractStrings(sub));
                                        });
                                    } else if (typeof obj === 'object' && obj !== null) {
                                        for (const key in obj) {
                                            result = result.concat(extractStrings(obj[key]));
                                        }
                                    }
                                    return result;
                                }
                                extractStrings(item).forEach(str => {
                                    if (str && str.trim()) lines.push(str);
                                });
                            }
                        });
                        translation = lines.join('\n');
                    }
                    // 例句
                    let exampleList: any[] = [];
                    function stripHtml(str: string) {
                        return str ? str.replace(/<[^>]+>/g, '').trim() : '';
                    }
                    if (w && w.exam_sents && w.exam_sents.length > 0) {
                        exampleList = w.exam_sents
                            .filter((s: any) => s.eng && s.chn)
                            .map((s: any) => `${stripHtml(s.eng)}\n${stripHtml(s.chn)}`);
                    }
                    if (dictData.media_sents_part && Array.isArray(dictData.media_sents_part.sent) && dictData.media_sents_part.sent.length > 0) {
                        exampleList = exampleList.concat(
                            dictData.media_sents_part.sent
                                .filter((s: any) => s.eng && s.chn)
                                .map((s: any) => `${stripHtml(s.eng)}\n${stripHtml(s.chn)}`)
                        );
                    }
                    if (dictData.blc && dictData.blc.blc_sents && dictData.blc.blc_sents.length > 0) {
                        exampleList = exampleList.concat(
                            dictData.blc.blc_sents
                                .filter((s: any) => s.eng && s.chn)
                                .map((s: any) => `${stripHtml(s.eng)}\n${stripHtml(s.chn)}`)
                        );
                    }
                    if (dictData.collins && dictData.collins.collins_entries && dictData.collins.collins_entries.length > 0) {
                        dictData.collins.collins_entries.forEach((collinsEntry: any) => {
                            if (collinsEntry.entries && collinsEntry.entries.entry) {
                                collinsEntry.entries.entry.forEach((entry: any) => {
                                    if (entry.tran_entry) {
                                        entry.tran_entry.forEach((tran: any) => {
                                            if (tran.exam_sents && tran.exam_sents.sent) {
                                                tran.exam_sents.sent.forEach((sent: any) => {
                                                    const eng = sent.eng_sent || sent.eng || '';
                                                    const chn = sent.chn_sent || sent.chn || '';
                                                    if (eng && chn) {
                                                        exampleList.push(`${stripHtml(eng)}\n${stripHtml(chn)}`);
                                                    }
                                                });
                                            }
                                        });
                                    }
                                });
                            }
                        });
                    }
                    if (exampleList.length > 0) {
                        example = exampleList.join('\n\n');
                    }
                    // 网络释义
                    if (dictData.web_trans && dictData.web_trans['web-translation'] && dictData.web_trans['web-translation'].length > 0) {
                        notes = dictData.web_trans['web-translation'].map((item: any) => {
                            return `${item.key}: ${(item.trans || []).map((t: any) => t.value).join('; ')}`;
                        }).join('\n');
                    }
                }
                // 插入到词汇本最前面
                const newItem = {
                    word,
                    translation,
                    example,
                    group: '',
                    notes,
                    mastered: false,
                    addedAt: Date.now()
                };
                this.vocabData.vocabBook.unshift(newItem);
                setTimeout(() => { this.saveVocabData(); }, 0);
                new Notice('已自动加入词汇本');
            }
        });

        // === 新增：打开词汇本快捷键命令 ===
        this.addCommand({
            id: "youdao-open-vocab-book",
            name: "打开词汇本（最大化并聚焦）",
            hotkeys: [],
            callback: () => {
                if (globalVocabBookModal) {
                    globalVocabBookModal._isMaximized = true;
                    globalVocabBookModal.contentEl.style.display = '';
                    globalVocabBookModal.modalEl.style.display = '';
                    globalVocabBookModal.onOpen();
                } else {
                    globalVocabBookModal = new VocabBookModal(this.app, this);
                    globalVocabBookModal._isMaximized = true;
                    globalVocabBookModal.open();
                }
            }
        });

        // === 新增：始终机器翻译选中文本命令 ===
        this.addCommand({
            id: "youdao-force-machine-translate-selection",
            name: "始终机器翻译选中文本（弹窗显示）",
            editorCallback: async (editor: Editor) => {
                if (hasMinimizedModal()) {
                    new Notice("请先还原/关闭顶层弹窗，再进行其他操作");
                    return;
                }
                const selectedText = (window.getSelection()?.toString() || '') || editor.getSelection();
                if (!selectedText) {
                    new Notice("请先选中要翻译的内容");
                    return;
                }
                // 多行/多句分割，逐行翻译
                const lines = selectedText.split(/\r?\n/).filter(line => line.trim().length > 0);
                // 判断目标语言函数
                const getTargetLang = (text: string) => {
                    if (this.settings.isBiDirection) {
                        return /[\u4e00-\u9fa5]/.test(text)
                            ? (this.settings.translateModel === 'composite' ? 'en' : 'en')
                            : (this.settings.translateModel === 'composite'
                                ? (this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang)
                                : (this.settings.textTargetLang === 'zh' ? 'zh-CHS' : this.settings.textTargetLang));
                    } else {
                        return this.settings.translateModel === 'composite'
                            ? (this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang)
                            : (this.settings.textTargetLang === 'zh' ? 'zh-CHS' : this.settings.textTargetLang);
                    }
                };
                const color = this.settings.textTranslationColor;
                const mode = this.settings.textTranslationMode;
                if (this.settings.translateModel === 'composite') {
                    // 复合模式：微软机器翻译，分行批量
                    await unifiedBatchTranslateForMicrosoft({
                        app: this.app,
                        items: lines,
                        getTargetLang: this.settings.isBiDirection
                            ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans'
                            : () => this.settings.textTargetLang === 'zh' ? 'zh-Hans' : this.settings.textTargetLang,
                        msAdapter: new MicrosoftAdapter(
                            this.settings.microsoftKey || '',
                            this.settings.microsoftRegion || '',
                            this.settings.microsoftEndpoint || '',
                            this.app
                        ),
                        color,
                        mode,
                        sleepInterval: this.settings.microsoftSleepInterval ?? 250,
                        plugin: this
                    });
                } else {
                    // 有道机器翻译，分行批量
                    await unifiedBatchTranslate({
                        app: this.app,
                        items: lines,
                        getTargetLang: this.settings.isBiDirection
                            ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-CHS'
                            : () => this.settings.textTargetLang,
                        textAppKey: this.settings.textAppKey,
                        textAppSecret: this.settings.textAppSecret,
                        color,
                        mode,
                        port: this.settings.serverPort,
                        sleepInterval: this.settings.sleepInterval ?? 250,
                        plugin: this,
                        isBiDirection: this.settings.isBiDirection
                    });
                }
            }
        });
    }

    // 只用 data 文件夹下的 json 文件，不再用主 data.json
    async loadAllData() {
        console.log('[YoudaoPlugin] 开始加载所有数据...');
        await ensureDataDir(this.app);
        this.settings = Object.assign({}, DEFAULT_SETTINGS, await this.readJsonFile('settings.json'));
        this.vocabData = Object.assign({}, DEFAULT_VOCAB_DATA, await this.readJsonFile('vocabBook.json'));
        this.trashData = Object.assign({}, DEFAULT_TRASH_DATA, await this.readJsonFile('vocabTrash.json'));
        this.serverHistory = Object.assign({}, DEFAULT_SERVER_HISTORY, await this.readJsonFile('serverHistory.json'));
        console.log('[YoudaoPlugin] 数据加载完成 - 词汇本:', this.vocabData.vocabBook?.length || 0, '回收站:', this.trashData.vocabBookTrash?.length || 0, '服务器历史:', this.serverHistory.serverPathHistory?.length || 0);
    }
    async saveSettings() {
        await ensureDataDir(this.app);
        console.log('[YoudaoPlugin] 保存设置数据:', this.settings);
        await this.writeJsonFile('settings.json', this.settings);
    }
    async saveVocabData() {
        await ensureDataDir(this.app);
        console.log('[YoudaoPlugin] 保存词汇本数据，词条数量:', this.vocabData.vocabBook?.length || 0);
        await this.writeJsonFile('vocabBook.json', this.vocabData);
    }
    async saveTrashData() {
        await ensureDataDir(this.app);
        console.log('[YoudaoPlugin] 保存回收站数据，词条数量:', this.trashData.vocabBookTrash?.length || 0);
        await this.writeJsonFile('vocabTrash.json', this.trashData);
    }
    async saveServerHistory() {
        await ensureDataDir(this.app);
        console.log('[YoudaoPlugin] 保存服务器历史数据，历史记录数量:', this.serverHistory.serverPathHistory?.length || 0);
        await this.writeJsonFile('serverHistory.json', this.serverHistory);
    }
    async readJsonFile(filename: string) {
        const relPath = `.obsidian/plugins/Obsidian Translation/data/${filename}`;
        try {
            console.log('[YoudaoPlugin] 读取文件:', relPath);
            const content = await this.app.vault.adapter.read(relPath);
            const data = JSON.parse(content);
            console.log('[YoudaoPlugin] 成功读取文件:', filename, '数据:', data);
            return data;
        } catch (e) {
            console.log('[YoudaoPlugin] 读取文件失败:', filename, '错误:', e);
            return {};
        }
    }
    async writeJsonFile(filename: string, data: any) {
        const relPath = `.obsidian/plugins/Obsidian Translation/data/${filename}`;
        console.log('[YoudaoPlugin] 写入文件:', relPath);
        try {
            await this.app.vault.adapter.write(relPath, JSON.stringify(data, null, 2));
            console.log('[YoudaoPlugin] 成功写入文件:', filename);
        } catch (e) {
            console.error('[YoudaoPlugin] 写入文件失败:', filename, '错误:', e);
            throw e;
        }
    }


    async startServer(isSwitchingModeParam = false) {
        const port = parseInt(this.settings.serverPort) || 4000;
        const translateModel = this.settings.translateModel || 'youdao';
        let targetPath = '';
        let isCompositeMode = false;
        if (translateModel === 'composite') {
            const dictServicePath = this.settings.dictServicePath || '';
            const path = require('path');
            targetPath = dictServicePath ? path.join(dictServicePath, 'api_dict.exe') : '';
            isCompositeMode = true;
        } else {
            targetPath = this.settings.serverPath;
            isCompositeMode = false;
        }
        const path = require('path');
        const serverDir = path.dirname(targetPath);
        let spawnCmd = '';
        if (isCompositeMode) {
            spawnCmd = `"${targetPath}" --port ${port}`;
        } else {
            spawnCmd = `"${targetPath}" ${port}`;
        }
        let started = false, exited = false, noticeShown = false;
        console.log('[复合模式] startServer参数:', {
            isSwitchingModeParam,
            translateModel,
            isCompositeMode,
            targetPath,
            port,
            spawnCmd
        });
        if (!isSwitchingModeParam) new Notice("正在启动中转服务...");
        // if (typeof this._refreshStatus === 'function') this._refreshStatus(); // <-- 注释掉，避免覆盖"正在开启中转服务..."
        console.log('[复合模式] startServer: isSwitchingModeParam=', isSwitchingModeParam, 'cmd=', spawnCmd, 'port=', port, 'mode=', translateModel);
        if (this.serverProcess && this._currentServerPort === port && this._currentServerPath === targetPath) {
            console.log('[复合模式] startServer: 端口已被占用，直接返回', { port, targetPath });
            new Notice(`端口${port}已被占用，可能服务已在运行或被其他程序占用`);
            if (typeof this._refreshStatus === 'function') this._refreshStatus();
            return;
        }
        if (this.serverProcess) {
            if (!isSwitchingModeParam) new Notice("正在关闭已有中转服务，准备切换...");
            if (typeof this._refreshStatus === 'function') this._refreshStatus();
            console.log('[复合模式] startServer: 先关闭已有服务');
            await this.stopServer(true);
            await waitPortRelease(port);
        }
        if (!targetPath) {
            const pathType = isCompositeMode ? 'api_dict.exe' : 'server.exe';
            new Notice(`请先设置${pathType}的路径`);
            if (typeof this._refreshStatus === 'function') this._refreshStatus();
            console.log('[复合模式] startServer: 未设置 targetPath');
            return;
        }
        const isPortInUse = await checkPortInUse(port);
        console.log('[复合模式] startServer: 检查端口是否被占用', { port, isPortInUse });
        if (isPortInUse) {
            new Notice(`端口${port}已被占用，可能服务已在运行或被其他程序占用`);
            if (typeof this._refreshStatus === 'function') this._refreshStatus();
            console.log('[复合模式] startServer: 端口被占用');
            return;
        }
        try {
            const { spawn } = require('child_process');
            this.serverProcess = spawn(
                spawnCmd,
                [],
                {
                cwd: serverDir,
                    stdio: 'pipe',
                    shell: true,
                    env: process.env
                }
            );
            console.log('[复合模式] startServer: 已spawn进程', { spawnCmd, serverDir });
            this.serverProcess.stdout.on('data', (data: Buffer) => {
                const msg = data.toString();
                console.log('[复合模式] startServer: stdout:', msg);
                if (
                    msg.includes('中转服务器已启动') ||
                    msg.includes('Server started') ||
                    msg.includes('Listening on port') ||
                    msg.includes('Running on http://127.0.0.1:') // 兼容 Flask
                ) {
                    started = true;
                    this._currentServerPort = port;
                    this._currentServerPath = targetPath;
                    const serviceType = isCompositeMode ? 'api_dict.exe' : 'server.exe';
                    if (!isSwitchingModeParam) new Notice(`${serviceType}中转服务启动成功`);
                    if (typeof this._refreshStatus === 'function') this._refreshStatus();
                    console.log('[复合模式] startServer: 服务启动成功，端口=', port, 'serviceType=', serviceType);
                }
            });
            this.serverProcess.stderr.on('data', (data: Buffer) => {
                const msg = data.toString();
                console.error('[复合模式] 服务器错误:', msg);
                console.log('[复合模式] startServer: stderr:', msg);
                // 新增：stderr 也检测 Flask 端口信息
                if (msg.includes('Running on http://127.0.0.1:')) {
                    started = true;
                    this._currentServerPort = port;
                    this._currentServerPath = targetPath;
                    const serviceType = isCompositeMode ? 'api_dict.exe' : 'server.exe';
                    if (!isSwitchingModeParam) new Notice(`${serviceType}中转服务启动成功`);
                    if (typeof this._refreshStatus === 'function') this._refreshStatus();
                    console.log('[复合模式] startServer: 服务启动成功（stderr），端口=', port, 'serviceType=', serviceType);
                }
            });
            this.serverProcess.on('close', (code: number) => {
                console.log('[复合模式] startServer: 服务器进程退出，代码:', code);
                this.serverProcess = null;
                exited = true;
                if (!started) {
                    const serviceType = isCompositeMode ? 'api_dict.exe' : 'server.exe';
                    if (!isSwitchingModeParam) new Notice(`${serviceType}中转服务启动失败，请检查路径和依赖，若稍后弹出'启动成功'则无需理会本提示`);
                    console.log('[复合模式] startServer: 服务启动失败');
                }
                if (typeof this._refreshStatus === 'function') this._refreshStatus();
            });
            this.serverProcess.on('error', (err: any) => {
                console.error('[复合模式] 服务器进程启动异常:', err);
                console.log('[复合模式] startServer: 进程启动异常', err);
                if (typeof this._refreshStatus === 'function') this._refreshStatus();
            });
            const timeoutMs = (this.settings.serverPathMode === 'cloud') ? 3000 : 1000;
            setTimeout(() => {
                if (!started && !noticeShown) {
                    const serviceType = isCompositeMode ? 'api_dict.exe' : 'server.exe';
                    if (!isSwitchingModeParam) new Notice(`${serviceType}中转服务启动失败，请检查路径和依赖，若稍后弹出'启动成功'则无需理会本提示`);
                    noticeShown = true;
                    if (this.serverProcess) {
                        this.serverProcess.kill('SIGINT');
                        this.serverProcess = null;
                        console.log('[复合模式] startServer: 超时未启动成功，强制kill');
                    }
                }
            }, timeoutMs);
        } catch (error) {
            console.error('[复合模式] 启动服务器失败:', error);
            const serviceType = isCompositeMode ? 'api_dict.exe' : 'server.exe';
            if (!isSwitchingModeParam) new Notice(`${serviceType}中转服务启动失败: ` + error.message + (error.stack ? '\n' + error.stack : ''));
            if (typeof this._refreshStatus === 'function') this._refreshStatus();
            console.log('[复合模式] startServer: catch error', error);
        }
    }

    async stopServer(isSwitching = false): Promise<void> {
        const port = parseInt(this.settings.serverPort) || 4000;
        const platform = process.platform;
        const translateModel = this.settings.translateModel || 'youdao';
        const serviceType = translateModel === 'composite' ? 'api_dict.exe' : 'server.exe';
        
        // 立即提示
        if (!isSwitching) {
            new Notice(`正在关闭${serviceType}中转端口${port}...`);
            if (typeof this._refreshStatus === 'function') this._refreshStatus();
        }
        // 检查端口是否被占用
        const isPortInUse = await checkPortInUse(port);
        console.log(`[stopServer] 检查端口${port}是否被占用:`, isPortInUse);
        if (!isPortInUse) {
            if (!isSwitching) new Notice(`${serviceType}中转端口${port}未开启`);
            if (typeof this._refreshStatus === 'function') this._refreshStatus();
            // 只有在端口真正关闭后才清空
            if (!isSwitching) {
                this._currentServerPort = undefined;
                this._currentServerPath = undefined;
            }
            return;
        }

        let findPidCmd = '';
        let killCmd = (pid: string) => '';
        if (platform === 'win32') {
            findPidCmd = getFindPidCmd(port);
            killCmd = (pid) => getKillCmd(pid);
        } else if (platform === 'darwin' || platform === 'linux') {
            findPidCmd = `lsof -i :${port} -t`;
            killCmd = (pid) => `kill -9 ${pid}`;
        } else {
            if (!isSwitching) new Notice('不支持的操作系统');
            return;
        }
        console.log(`[stopServer] 查找PID命令: ${findPidCmd}`);
        await new Promise<void>((resolve) => {
            exec(findPidCmd, (err: any, stdout: string, stderr: string) => {
                console.log(`[stopServer] 查找PID结果 err:`, err);
                console.log(`[stopServer] 查找PID结果 stdout:`, stdout);
                console.log(`[stopServer] 查找PID结果 stderr:`, stderr);
                let pids: string[] = [];
                if (!err && stdout) {
                    if (platform === 'win32') {
                        const lines = stdout.trim().split('\n');
                        const pidSet = new Set();
                        for (const line of lines) {
                            const parts = line.trim().split(/\s+/);
                            const pid = parts[parts.length - 1];
                            if (pid) pidSet.add(pid);
                        }
                        pids = Array.from(pidSet) as string[];
                    } else {
                        pids = stdout.trim().split('\n').filter(Boolean);
                    }
                }
                console.log(`[stopServer] 需要kill的PID:`, pids);
                if (pids.length === 0) {
                    if (!isSwitching) new Notice(`未找到占用端口${port}的${serviceType}进程`);
                    if (typeof this._refreshStatus === 'function') this._refreshStatus();
                    // 只有在端口真正关闭后才清空
                    if (!isSwitching) {
                        this._currentServerPort = undefined;
                        this._currentServerPath = undefined;
                    }
                    resolve();
                    return;
                }

                let killCount = 0;
                pids.forEach(pid => {
                    const cmd = killCmd(pid);
                    console.log(`[stopServer] 执行kill命令: ${cmd}`);
                    exec(cmd, async (killErr: any, killStdout: string, killStderr: string) => {
                        console.log(`[stopServer] kill结果 err:`, killErr);
                        console.log(`[stopServer] kill结果 stdout:`, killStdout);
                        console.log(`[stopServer] kill结果 stderr:`, killStderr);
                        killCount++;
                        // 新增：等待端口彻底释放，并多次尝试 kill
                        let released = false;
                        for (let i = 0; i < 20; i++) {
                            await sleep(100);
                            const inUse = await checkPortInUse(port);
                            console.log(`[stopServer] 等待端口释放: 第${i+1}次, inUse=`, inUse);
                            if (!inUse) {
                                released = true;
                                break;
                            }
                            // 新增：如果还在占用，每隔5次再尝试 kill 一次
                            if (inUse && i % 5 === 4) {
                                exec(cmd, () => {});
                            }
                        }
                        // 再判断端口是否还被占用
                        const stillInUse = await checkPortInUse(port);
                        console.log(`[stopServer] 最终端口占用状态:`, stillInUse);
                        if (stillInUse) {
                            if (!isSwitching) new Notice(`${serviceType}端口${port}关闭失败，端口仍被占用`);
                        } else {
                            if (!isSwitching) new Notice(`${serviceType}端口${port}关闭成功`);
                            if (!isSwitching) {
                                this._currentServerPort = undefined;
                                this._currentServerPath = undefined;
                            }
                        }
                        if (typeof this._refreshStatus === 'function') this._refreshStatus();
                        if (killCount === pids.length) resolve();
                    });
                });
            });
        });

        // 清理插件自己的 serverProcess 状态
        if (this.serverProcess) {
                this.serverProcess = null;
            // 不要提前清空 _currentServerPort
        }
    }

    // 获取 data 目录下的 json 文件路径（相对插件目录）
    getDataFilePath(filename: string): string {
        // 插件目录下 data 文件夹
        // @ts-ignore
        const base = (this.app.vault.adapter?.basePath || (this.app.vault.adapter?.getBasePath && this.app.vault.adapter.getBasePath()) || '.');
        const dataDir = path.join(base, '.obsidian', 'plugins', 'Obsidian Translation', 'data');
        if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir, { recursive: true });
        return path.join(dataDir, filename);
    }




}



function checkPortInUse(port: number): Promise<boolean> {
    const { exec } = require('child_process');
    const platform = process.platform;
    return new Promise((resolve) => {
        let cmd = '';
        if (platform === 'win32') {
            cmd = `netstat -ano | findstr :${port}`;
        } else {
            cmd = `lsof -i :${port} -t`;
        }
        exec(cmd, (err: any, stdout: string) => {
            resolve(!!stdout && stdout.trim().length > 0);
        });
    });
}

class YoudaoSettingTab extends PluginSettingTab {
    plugin: YoudaoTranslatePlugin;
    constructor(app: App, plugin: YoudaoTranslatePlugin) {
        super(app, plugin);
        this.plugin = plugin;
    }
    display(): void {
        const { containerEl } = this;
        containerEl.empty();
        let inputEl: HTMLInputElement; // <--- 这里声明
        // 新增：中英互译开关
        let isBiDirection = this.plugin.settings.isBiDirection || false;
        const biDirectionDiv = containerEl.createEl("div");
        biDirectionDiv.style.margin = "16px 0";
        biDirectionDiv.style.display = "flex";
        biDirectionDiv.style.alignItems = "center";
        biDirectionDiv.style.gap = "8px";
        const biDirectionLabel = biDirectionDiv.createEl("span", { text: "中英互译（自动判断目标语言）" });
        const biDirectionBtn = biDirectionDiv.createEl("button", { text: isBiDirection ? "已开启" : "未开启" });
        biDirectionBtn.style.background = isBiDirection ? "#1a73e8" : "#eee";
        biDirectionBtn.style.color = isBiDirection ? "#fff" : "#333";
        biDirectionBtn.style.border = "1px solid #ccc";
        biDirectionBtn.style.borderRadius = "4px";
        biDirectionBtn.style.padding = "2px 12px";
        biDirectionBtn.style.cursor = "pointer";
        biDirectionBtn.onclick = async () => {
            isBiDirection = !isBiDirection;
            this.plugin.settings.isBiDirection = isBiDirection;
            await this.plugin.saveSettings();
            biDirectionBtn.textContent = isBiDirection ? "已开启" : "未开启";
            biDirectionBtn.style.background = isBiDirection ? "#1a73e8" : "#eee";
            biDirectionBtn.style.color = isBiDirection ? "#fff" : "#333";
            // 触发界面刷新以禁用/启用下拉框
            this.display();
        };
        containerEl.appendChild(biDirectionDiv);
        const howToUseBtn = containerEl.createEl("button", { text: "如何使用插件" });
        // 新增：词典释义/机器翻译分界单词数输入框
        const dictLimitDiv = containerEl.createEl("div");
        dictLimitDiv.style.margin = "16px 0";
        dictLimitDiv.style.display = "flex";
        dictLimitDiv.style.alignItems = "center";
        dictLimitDiv.style.gap = "8px";
        dictLimitDiv.createEl("span", { text: "词典释义/机器翻译分界（英文单词数/中文汉字数）：" });
        // 英文分界
        const dictLimitInputEn = dictLimitDiv.createEl("input");
        dictLimitInputEn.type = "number";
        dictLimitInputEn.min = "1";
        dictLimitInputEn.max = "20";
        dictLimitInputEn.style.width = "40px";
        dictLimitInputEn.value = String(this.plugin.settings.dictWordLimitEn ?? 3);
        dictLimitInputEn.onchange = async () => {
            let val = parseInt(dictLimitInputEn.value);
            if (isNaN(val) || val < 1) val = 1;
            if (val > 20) val = 20;
            dictLimitInputEn.value = String(val);
            this.plugin.settings.dictWordLimitEn = val;
            await this.plugin.saveSettings();
        };
        dictLimitDiv.createEl("span", { text: "英文" });
        // 中文分界
        const dictLimitInputZh = dictLimitDiv.createEl("input");
        dictLimitInputZh.type = "number";
        dictLimitInputZh.min = "1";
        dictLimitInputZh.max = "20";
        dictLimitInputZh.style.width = "40px";
        dictLimitInputZh.value = String(this.plugin.settings.dictWordLimitZh ?? 4);
        dictLimitInputZh.onchange = async () => {
            let val = parseInt(dictLimitInputZh.value);
            if (isNaN(val) || val < 1) val = 1;
            if (val > 20) val = 20;
            dictLimitInputZh.value = String(val);
            this.plugin.settings.dictWordLimitZh = val;
            await this.plugin.saveSettings();
        };
        dictLimitDiv.createEl("span", { text: "中文" });
        containerEl.appendChild(dictLimitDiv);
        // datalist相关代码移到这里，避免变量未声明报错
        const history = this.plugin.serverHistory.serverPathHistory || [];
        const datalistId = "server-path-history-list";
        let datalist = containerEl.querySelector(`#${datalistId}`) as HTMLDataListElement;
        if (datalist) datalist.remove();
        datalist = document.createElement("datalist");
        datalist.id = datalistId;
        history.forEach(p => {
            const option = document.createElement("option");
            option.value = p;
            datalist.appendChild(option);
        });
        containerEl.appendChild(datalist);
        howToUseBtn.style.margin = "16px 0";
        howToUseBtn.onclick = () => {
            new HowToUseModal(this.app).open();
        };
        // --- 中转服务器设置标题和清理日志按钮 ---
        const serverHeaderRow = containerEl.createEl("div");
        serverHeaderRow.style.display = "flex";
        serverHeaderRow.style.alignItems = "center";
        serverHeaderRow.style.justifyContent = "space-between";
        serverHeaderRow.style.marginTop = "24px";

        // 左侧标题
        const serverHeader = serverHeaderRow.createEl("h2", { text: "中转服务器设置" });

        // 右侧清理日志按钮
        const clearLogBtn = serverHeaderRow.createEl("button", { text: "清理日志" });
        clearLogBtn.style.fontSize = "14px";
        clearLogBtn.style.padding = "4px 12px";
        clearLogBtn.style.cursor = "pointer";
        clearLogBtn.onclick = async () => {
            const fs = require('fs');
            const path = require('path');
            const serverPath = this.plugin.settings.serverPath;
            if (!serverPath) {
                new Notice("请先设置 server.exe 绝对路径");
                return;
            }
            const serverDir = path.dirname(serverPath);
            const logPath = path.join(serverDir, 'start.log');
            try {
                if (fs.existsSync(logPath)) {
                    fs.unlinkSync(logPath);
                    new Notice("start.log 日志已清理");
                } else {
                    new Notice("start.log 文件不存在");
                }
            } catch (e) {
                new Notice("清理日志失败: " + e.message);
            }
        };
        // 新增：清理所有api_dict.exe进程按钮
        const clearApiDictBtn = serverHeaderRow.createEl("button", { text: "清理所有api_dict.exe进程" });
        clearApiDictBtn.style.fontSize = "14px";
        clearApiDictBtn.style.padding = "4px 12px";
        clearApiDictBtn.style.cursor = "pointer";
        clearApiDictBtn.style.marginLeft = "12px";
        clearApiDictBtn.onclick = async () => {
            const { exec } = require('child_process');
            const platform = process.platform;
            let cmd = '';
            if (platform === 'win32') {
                cmd = 'taskkill /IM api_dict.exe /F';
            } else {
                cmd = 'pkill -f api_dict.exe';
            }
            exec(cmd, { encoding: 'buffer' }, (err: any, stdout: Buffer, stderr: Buffer) => {
                let msg = '';
                if (platform === 'win32') {
                    // 尝试用 GBK 解码
                    const iconv = require('iconv-lite');
                    msg = iconv.decode(stderr, 'gbk');
                    if (msg.includes('未找到') || msg.includes('not found')) {
                        new Notice('所有 api_dict.exe 进程已关闭');
                        return;
                    }
                }
                if (err) {
                    new Notice('清理失败: ' + (msg || err.message));
                } else {
                    new Notice('已清理所有 api_dict.exe 进程');
                }
            });
        };
        // 新增：清理所有server.exe进程按钮
        const clearServerExeBtn = serverHeaderRow.createEl("button", { text: "清理所有server.exe进程" });
        clearServerExeBtn.style.fontSize = "14px";
        clearServerExeBtn.style.padding = "4px 12px";
        clearServerExeBtn.style.cursor = "pointer";
        clearServerExeBtn.style.marginLeft = "12px";
        clearServerExeBtn.onclick = async () => {
            const { exec } = require('child_process');
            const platform = process.platform;
            let cmd = '';
            if (platform === 'win32') {
                cmd = 'taskkill /IM server.exe /F';
            } else {
                cmd = 'pkill -f server.exe';
            }
            exec(cmd, { encoding: 'buffer' }, (err: any, stdout: Buffer, stderr: Buffer) => {
                let msg = '';
                if (platform === 'win32') {
                    // 尝试用 GBK 解码
                    const iconv = require('iconv-lite');
                    msg = iconv.decode(stderr, 'gbk');
                    if (msg.includes('未找到') || msg.includes('not found')) {
                        new Notice('所有 server.exe 进程已关闭');
                        return;
                    }
                }
                if (err) {
                    new Notice('清理失败: ' + (msg || err.message));
                } else {
                    new Notice('已清理所有 server.exe 进程');
                }
            });
        };

        containerEl.createEl("h2", { text: "中转服务器设置" });
        // server.exe 路径历史补全
        const serverJsDatalistId = "server-js-history";
        const serverJsDatalist = document.createElement("datalist");
        serverJsDatalist.id = serverJsDatalistId;
        (containerEl as HTMLElement).appendChild(serverJsDatalist);
        let serverJsInputEl: HTMLInputElement | null = null;
        new Setting(containerEl)
            .setName("server.exe 路径")
            .setDesc("请填写 server.exe 的绝对路径,填写完后请先手动运行server.exe(避免出错）")
            .addText(text => {
                text.setPlaceholder("如 D:/xxx/server.exe")
                    .setValue(this.plugin.settings.serverPath)
                    .onChange(async (value) => {
                        this.plugin.settings.serverPath = value;
                        // 不自动保存历史，只有点击保存按钮才保存
                        await this.plugin.saveSettings();
                    });
                text.inputEl.setAttribute("list", serverJsDatalistId);
                serverJsInputEl = text.inputEl;
                // 初始化 datalist
                const history = this.plugin.serverHistory.serverPathHistory || [];
                while (serverJsDatalist.firstChild) serverJsDatalist.removeChild(serverJsDatalist.firstChild);
                history.forEach(p => {
                    const option = document.createElement("option");
                    option.value = p;
                    serverJsDatalist.appendChild(option);
                });
            })
            .addButton(btn => btn
                .setButtonText("保存")
                .onClick(async () => {
                    const value = serverJsInputEl?.value.trim() || '';
                    if (!value) {
                        new Notice("请输入地址后保存");
                        return;
                    }
                    // 保存到历史
                    let history = this.plugin.serverHistory.serverPathHistory || [];
                    history = [value, ...history.filter(p => p !== value)];
                    if (history.length > 10) history = history.slice(0, 10);
                    this.plugin.serverHistory.serverPathHistory = history;
                    this.plugin.settings.serverPath = value;
                    await this.plugin.saveServerHistory();
                    await this.plugin.saveSettings();
                    // 更新 datalist
                    while (serverJsDatalist.firstChild) serverJsDatalist.removeChild(serverJsDatalist.firstChild);
                    history.forEach(p => {
                        const option = document.createElement("option");
                        option.value = p;
                        serverJsDatalist.appendChild(option);
                    });
                    new Notice("地址保存成功");
                }));

         // dict_service 路径设置项，紧跟在api_dict.exe路径下方
         const dictServiceDatalistId = "dict-service-path-history";
         const dictServiceDatalist = document.createElement("datalist");
         dictServiceDatalist.id = dictServiceDatalistId;
         (containerEl as HTMLElement).appendChild(dictServiceDatalist);
         let dictServiceInputEl: HTMLInputElement | null = null;
         new Setting(containerEl)
             .setName("dict_service 路径")
             .setDesc("请填写 dict_service 文件夹的绝对路径")
             .addText(text => {
                 text.setPlaceholder("如 D:/xxx/dict_service")
                     .setValue(this.plugin.settings.dictServicePath || "")
                     .onChange(async (value) => {
                         this.plugin.settings.dictServicePath = value;
                         await this.plugin.saveSettings();
                     });
                 text.inputEl.setAttribute("list", dictServiceDatalistId);
                 dictServiceInputEl = text.inputEl;
                 // 初始化 datalist
                 const history = this.plugin.serverHistory.dictServicePathHistory || [];
                 while (dictServiceDatalist.firstChild) dictServiceDatalist.removeChild(dictServiceDatalist.firstChild);
                 history.forEach(p => {
                     const option = document.createElement("option");
                     option.value = p;
                     dictServiceDatalist.appendChild(option);
                 });
             })
             .addButton(btn => btn
                 .setButtonText("保存")
                 .onClick(async () => {
                     const value = dictServiceInputEl?.value.trim() || '';
                     if (!value) {
                         new Notice("请输入地址后保存");
                         return;
                     }
                     // 路径存在性提示（不阻止保存）
                     try {
                         // @ts-ignore
                         const exists = await this.plugin.app.vault.adapter.exists(value);
                         if (!exists) {
                             new Notice("注意：该路径在本地未检测到，但仍已保存");
                         }
                     } catch (e) {
                         // 忽略异常
                     }
                     // 保存到历史
                     let history = this.plugin.serverHistory.dictServicePathHistory || [];
                     history = [value, ...history.filter(p => p !== value)];
                     if (history.length > 10) history = history.slice(0, 10);
                     this.plugin.serverHistory.dictServicePathHistory = history;
                     this.plugin.settings.dictServicePath = value;
                     await this.plugin.saveServerHistory();
                     await this.plugin.saveSettings();
                     // 更新 datalist
                     while (dictServiceDatalist.firstChild) dictServiceDatalist.removeChild(dictServiceDatalist.firstChild);
                     history.forEach(p => {
                         const option = document.createElement("option");
                         option.value = p;
                         dictServiceDatalist.appendChild(option);
                     });
                     new Notice("地址保存成功");
                 }));

        new Setting(containerEl)
            .setName("server.exe 路径类型")
            .setDesc("请选择 server.exe 所在磁盘类型，影响服务启动等待时间")
            .addDropdown(drop => {
                drop.addOption('local', '本地磁盘');
                drop.addOption('cloud', '云盘/U盘');
                drop.setValue(this.plugin.settings.serverPathMode || 'local');
                drop.onChange(async (value: 'local' | 'cloud') => {
                    this.plugin.settings.serverPathMode = value;
                    await this.plugin.saveSettings();
                });
            });

    
    
        new Setting(containerEl)
            .setName("中转服务端口")
            .setDesc("如 4000，需与 server.exe 保持一致")
            .addText(text => text
                .setPlaceholder("4000")
                .setValue(this.plugin.settings.serverPort)
                .onChange(async (value) => {
                    this.plugin.settings.serverPort = value;
                    await this.plugin.saveSettings();
                }));

        // 新增：中转服务器控制
        new Setting(containerEl)
            .setName("中转服务器控制")
            .setDesc("启动或关闭本地中转服务器")
            .addButton(btn => btn
                .setButtonText("开启中转服务")
                .onClick(async () => {
                    // 根据翻译模式选择路径
                    const mode = this.plugin.settings.translateModel;
                    let targetPath = '';
                    if (mode === 'composite') {
                        const dictServicePath = this.plugin.settings.dictServicePath || '';
                        const path = require('path');
                        targetPath = dictServicePath ? path.join(dictServicePath, 'api_dict.exe') : '';
                    } else {
                        targetPath = this.plugin.settings.serverPath;
                    }
                    if (!targetPath) {
                        new Notice("请先设置对应的服务路径");
                        return;
                    }
                    // 通过插件方法设置 _currentServerPath
                    (this.plugin as any)._currentServerPath = targetPath;
                    await this.plugin.startServer();
                }))
            .addButton(btn => btn
                .setButtonText("关闭中转服务")
                .onClick(async () => {
                    // 根据翻译模式选择路径
                    const mode = this.plugin.settings.translateModel;
                    let targetPath = '';
                    if (mode === 'composite') {
                        const dictServicePath = this.plugin.settings.dictServicePath || '';
                        const path = require('path');
                        targetPath = dictServicePath ? path.join(dictServicePath, 'api_dict.exe') : '';
                    } else {
                        targetPath = this.plugin.settings.serverPath;
                    }
                    if (!targetPath) {
                        new Notice("请先设置对应的服务路径");
                        return;
                    }
                    (this.plugin as any)._currentServerPath = targetPath;
                    await this.plugin.stopServer();
                }));

        // --- 新增显示框 ---
        const statusDiv = containerEl.createEl("div");
        statusDiv.style.margin = "12px 0";
        statusDiv.style.fontWeight = "bold";
        statusDiv.style.color = "#1a73e8";

        // 状态管理变量
        let isSwitchingMode = false;
        let switchingPromise: Promise<void> | null = null;

        // 统一蓝色字体刷新逻辑，内容根据 mode 区分
        const refreshStatus = (msg?: string) => {
            if (isSwitchingMode) {
                statusDiv.textContent = "正在切换...";
                return;
            }
            if (msg) {
                statusDiv.textContent = msg;
                return;
            }
            const mode = this.plugin.settings.translateModel;
            const port = (this.plugin as any)._currentServerPort;
            if (port) {
                if (mode === 'composite') {
                    statusDiv.textContent = `当前复合翻译端口：${port}`;
                } else {
                    statusDiv.textContent = `当前有道中转服务端口：${port}`;
                }
            } else {
                statusDiv.textContent = "";
            }
        };
        refreshStatus();

        // 按钮点击时先显示"正在启动/关闭..."
        const openBtn = containerEl.querySelector("button:nth-of-type(1)");
        const closeBtn = containerEl.querySelector("button:nth-of-type(2)");
        if (openBtn) openBtn.addEventListener("click", () => {
            isSwitchingMode = false;
            console.log('[调试] 点击开启中转服务按钮');
            refreshStatus("正在开启中转服务端口...");
        });
        if (closeBtn) closeBtn.addEventListener("click", () => {
            isSwitchingMode = false;
            console.log('[调试] 点击关闭中转服务按钮');
            refreshStatus("正在关闭中转服务...");
        });

        // 监听模式切换，显示"正在切换..."
        let lastTranslateModel = this.plugin.settings.translateModel;
        const origSetValue = (Setting.prototype as any).setValue;
        (Setting.prototype as any).setValue = function(value: any) {
            if (
                this.settingName === 'server.exe 路径类型' ||
                this.settingName === 'api_dict.exe 路径' ||
                this.settingName === 'server.exe 路径' ||
                this.settingName === 'api_dict.exe 路径' ||
                this.settingName === '翻译模式'
            ) {
                isSwitchingMode = true;
                refreshStatus("正在切换...");
                const plugin = (this.plugin as any);
                const oldModel = lastTranslateModel;
                const oldPort = plugin.settings.serverPort;
                (async () => {
                    try {
                        console.log(`[切换模式] 旧模式: ${oldModel}, 旧端口: ${oldPort}, 新值:`, value);
                        if (oldModel === 'composite' && value === 'youdao') {
                            let basePort = parseInt(oldPort) || 4000;
                            let newPort = await findAvailablePort(basePort + 1, 20);
                            console.log(`[切换模式] composite->youdao，分配新端口: ${newPort}`);
                            plugin.settings.serverPort = String(newPort);
                            await plugin.saveSettings();
                            refreshStatus('正在切换（分配新端口 ' + newPort + '）...');
                            console.log('[切换模式] stopServer(true) for old composite');
                            await plugin.stopServer(true);
                            console.log('[切换模式] startServer(true) for new composite');
                            await plugin.startServer(true);
                            refreshStatus('新端口 ' + newPort + ' 已启动，继续切换...');
                        }
                        // 继续原有切换流程
                        console.log('[切换模式] stopServer(true) for target模式');
                        await plugin.stopServer(true);
                        console.log('[切换模式] startServer(true) for target模式');
                        await plugin.startServer(true);
                    } catch (e) {
                        console.error('[切换模式] 切换过程异常:', e);
                    } finally {
                        isSwitchingMode = false;
                        refreshStatus();
                    }
                })();
                lastTranslateModel = value;
            }
            if (origSetValue) return origSetValue.call(this, value);
        };

        // 暴露刷新方法给插件类
        (this.plugin as any)._refreshStatus = (msg?: string) => {
            // 切换结束条件：新端口启动成功
            if (isSwitchingMode && (this.plugin as any)._currentServerPort) {
                isSwitchingMode = false;
                console.log('[调试] _refreshStatus: 新端口启动成功, isSwitchingMode=false');
            }
            refreshStatus(msg);
        };

        containerEl.createEl("h2", { text: "文本翻译设置" });



        // new Setting(containerEl)
        //     .setName("文本 appKey")
        //     .setDesc("你的有道文本翻译应用ID")
        //     .addText(text => text
        //         .setPlaceholder("appKey")
        //         .setValue(this.plugin.settings.textAppKey)
        //         .onChange(async (value) => {
        //             this.plugin.settings.textAppKey = value;
        //             await this.plugin.saveSettings();
        //         }));
        // new Setting(containerEl)
        //     .setName("文本 appSecret")
        //     .setDesc("你的有道文本翻译应用密钥")
        //     .addText(text => text
        //         .setPlaceholder("appSecret")
        //         .setValue(this.plugin.settings.textAppSecret)
        //         .onChange(async (value) => {
        //             this.plugin.settings.textAppSecret = value;
        //             await this.plugin.saveSettings();
        //         }));
        // --- 设置界面目标语言下拉框 ---
        // 文本翻译目标语言
        new Setting(containerEl)
            .setName("文本翻译目标语言")
            .setDesc("选择翻译目标语言")
            .addDropdown(drop => drop
                .addOption("en", "英文")
                .addOption("zh-CHS", "中文")
                .setValue(this.plugin.settings.textTargetLang === 'zh' ? 'zh-CHS' : this.plugin.settings.textTargetLang)
                .onChange(async (value) => {
                    // 兜底：如果用户手动填了 zh，自动转为 zh-CHS
                    this.plugin.settings.textTargetLang = value === 'zh' ? 'zh-CHS' : value;
                    await this.plugin.saveSettings();
                })
                // 新增：中英互译时禁用
                .setDisabled(this.plugin.settings.isBiDirection === true)
            );
        new Setting(containerEl)
            .setName("文本译文颜色")
            .setDesc("设置文本翻译结果中译文的颜色")
            .addColorPicker(color => color
                .setValue(this.plugin.settings.textTranslationColor)
                .onChange(async (value) => {
                    this.plugin.settings.textTranslationColor = value;
                    await this.plugin.saveSettings();
                }));
        new Setting(containerEl)
            .setName("文本翻译模式")
            .setDesc("选择文本翻译结果的显示方式")
            .addDropdown(drop => drop
                .addOption("merge", "原句+译句交错显示")
                .addOption("split", "原句和译句分开显示")
                .setValue(this.plugin.settings.textTranslationMode)
                .onChange(async (value) => {
                    this.plugin.settings.textTranslationMode = value;
                    await this.plugin.saveSettings();
                }));

        containerEl.createEl("h2", { text: "图片翻译设置" });
        // new Setting(containerEl)
        //     .setName("图片 appKey")
        //     .setDesc("你的有道图片翻译应用ID")
        //     .addText(text => text
        //         .setPlaceholder("imageAppKey")
        //         .setValue(this.plugin.settings.imageAppKey)
        //         .onChange(async (value) => {
        //             this.plugin.settings.imageAppKey = value;
        //             await this.plugin.saveSettings();
        //         }));
        // new Setting(containerEl)
        //     .setName("图片 appSecret")
        //     .setDesc("你的有道图片翻译应用密钥")
        //     .addText(text => text
        //         .setPlaceholder("imageAppSecret")
        //         .setValue(this.plugin.settings.imageAppSecret)
        //         .onChange(async (value) => {
        //             this.plugin.settings.imageAppSecret = value;
        //             await this.plugin.saveSettings();
        //         }));
        // --- 设置界面目标语言下拉框 ---
        // 图片翻译目标语言
        new Setting(containerEl)
            .setName("图片翻译目标语言")
            .setDesc("选择图片翻译目标语言")
            .addDropdown(drop => drop
                .addOption("en", "英文")
                .addOption("zh-CHS", "中文")
                .setValue(this.plugin.settings.imageTargetLang === 'zh' ? 'zh-CHS' : this.plugin.settings.imageTargetLang)
                .onChange(async (value) => {
                    this.plugin.settings.imageTargetLang = value === 'zh' ? 'zh-CHS' : value;
                    await this.plugin.saveSettings();
                })
                // 新增：中英互译时禁用
                .setDisabled(this.plugin.settings.isBiDirection === true)
            );
        new Setting(containerEl)
            .setName("图片译文颜色")
            .setDesc("设置图片翻译结果中译文的颜色")
            .addColorPicker(color => color
                .setValue(this.plugin.settings.imageTranslationColor)
                .onChange(async (value) => {
                    this.plugin.settings.imageTranslationColor = value;
                    await this.plugin.saveSettings();
                }));
        new Setting(containerEl)
            .setName("图片翻译模式")
            .setDesc("选择图片翻译结果的显示方式")
            .addDropdown(drop => drop
                .addOption("merge", "原句+译句交错显示")
                .addOption("split", "原句和译句分开显示")
                .setValue(this.plugin.settings.imageTranslationMode)
                .onChange(async (value) => {
                    this.plugin.settings.imageTranslationMode = value;
                    await this.plugin.saveSettings();
                }));
        // 新增：多行文本/图片翻译速度设置
        // const speedDiv = containerEl.createEl("div");
        // speedDiv.style.margin = "16px 0";
        // speedDiv.style.display = "flex";
        // speedDiv.style.alignItems = "center";
        // speedDiv.style.gap = "8px";
        // speedDiv.createEl("span", { text: "多行文本/图片翻译速度（间隔ms，越小越快，API风控风险越高）:" });
        // const speedInput = speedDiv.createEl("input");
        // speedInput.type = "range";
        // speedInput.min = "50";
        // speedInput.max = "1000";
        // speedInput.step = "10";
        // speedInput.value = String(this.plugin.settings.sleepInterval ?? 250);
        // speedInput.style.width = "200px";
        // const speedVal = speedDiv.createEl("span", { text: speedInput.value });
        // speedInput.oninput = () => {
        //     speedVal.textContent = speedInput.value;
        // };
        // speedInput.onchange = async () => {
        //     this.plugin.settings.sleepInterval = parseInt(speedInput.value);
        //     await this.plugin.saveSettings();
        // };
        // containerEl.appendChild(speedDiv);

        // 词汇本按钮
        const vocabBtn = containerEl.createEl("button", { text: "打开词汇本" });
        vocabBtn.style.margin = "16px 0";
        vocabBtn.style.padding = "8px 24px";
        vocabBtn.style.fontSize = "1.1em";
        vocabBtn.style.background = "#1a73e8";
        vocabBtn.style.color = "#fff";
        vocabBtn.style.border = "none";
        vocabBtn.style.borderRadius = "6px";
        vocabBtn.style.cursor = "pointer";
        vocabBtn.onmouseenter = () => vocabBtn.style.background = "#1765c1";
        vocabBtn.onmouseleave = () => vocabBtn.style.background = "#1a73e8";
        vocabBtn.onclick = () => {
            if (globalVocabBookModal) {
                globalVocabBookModal.modalEl.style.display = '';
                globalVocabBookModal._isMinimized = false;
                globalVocabBookModal._isMaximized = true;
                globalVocabBookModal.modalEl.style.width = '98vw';
                globalVocabBookModal.modalEl.style.height = '96vh';
                globalVocabBookModal.modalEl.style.left = '1vw';
                globalVocabBookModal.modalEl.style.top = '2vh';
                globalVocabBookModal.modalEl.focus();
            } else {
                globalVocabBookModal = new VocabBookModal(this.app, this.plugin);
                globalVocabBookModal._isMaximized = true;
                globalVocabBookModal.open();
            }
        };
        containerEl.appendChild(vocabBtn);

        // Cursor 风格搜索区
        const searchBoxWrapper = containerEl.createEl("div");
        searchBoxWrapper.style.position = "absolute";
        searchBoxWrapper.style.top = "24px";
        searchBoxWrapper.style.right = "32px";
        searchBoxWrapper.style.zIndex = "1000";
        searchBoxWrapper.style.display = "flex";
        searchBoxWrapper.style.alignItems = "center";
        searchBoxWrapper.style.background = "#fff";
        searchBoxWrapper.style.border = "1px solid #e0e0e0";
        searchBoxWrapper.style.borderRadius = "8px";
        searchBoxWrapper.style.boxShadow = "0 2px 8px #0001";
        searchBoxWrapper.style.padding = "4px 8px";
        searchBoxWrapper.style.gap = "0";

        // 新增：搜索模式切换按钮
        let searchMode: 'fuzzy' | 'word' = 'fuzzy';
        let performSearch: () => void;
        let modeBtn: HTMLButtonElement;
        modeBtn = searchBoxWrapper.createEl("button", { text: "模糊搜索" });
        modeBtn.style.marginRight = "8px";
        modeBtn.style.padding = "4px 10px";
        modeBtn.style.fontSize = "0.98em";
        modeBtn.style.border = "none";
        modeBtn.style.background = "#e3eafc";
        modeBtn.style.borderRadius = "4px";
        modeBtn.style.cursor = "pointer";

        const searchInput = searchBoxWrapper.createEl("input");
        searchInput.type = "text";
        searchInput.placeholder = "输入单词搜索";
        searchInput.style.padding = "4px 8px";
        searchInput.style.fontSize = "1em";
        searchInput.style.border = "none";
        searchInput.style.outline = "none";
        searchInput.style.background = "transparent";

        const upBtn = searchBoxWrapper.createEl("button", { text: "↑" });
        upBtn.title = "上一个";
        upBtn.style.padding = "2px 8px";
        upBtn.style.margin = "0 2px";
        upBtn.style.border = "none";
        upBtn.style.background = "#f5f5f5";
        upBtn.style.borderRadius = "4px";
        upBtn.style.cursor = "pointer";

        const downBtn = searchBoxWrapper.createEl("button", { text: "↓" });
        downBtn.title = "下一个";
        downBtn.style.padding = "2px 8px";
        downBtn.style.margin = "0 2px";
        downBtn.style.border = "none";
        downBtn.style.background = "#f5f5f5";
        downBtn.style.borderRadius = "4px";
        downBtn.style.cursor = "pointer";

        const resultInfo = searchBoxWrapper.createEl("span");
        resultInfo.style.margin = "0 8px";
        resultInfo.style.fontSize = "0.98em";
        resultInfo.style.color = "#888";

        const searchBtn = searchBoxWrapper.createEl("button", { text: "搜索" });
        searchBtn.style.padding = "4px 12px";
        searchBtn.style.fontSize = "1em";
        searchBtn.style.marginLeft = "8px";
        searchBtn.style.border = "none";
        searchBtn.style.background = "#e3eafc";
        searchBtn.style.borderRadius = "4px";
        searchBtn.style.cursor = "pointer";

        containerEl.appendChild(searchBoxWrapper);

        // 在文本翻译设置前插入"有道翻译设置"按钮
        const youdaoConfigBtn = containerEl.createEl('button', { text: '有道翻译设置' });
        youdaoConfigBtn.style.margin = '16px 0';
        youdaoConfigBtn.style.padding = '8px 24px';
        youdaoConfigBtn.style.fontSize = '1.1em';
        youdaoConfigBtn.style.background = '#e67e22';
        youdaoConfigBtn.style.color = '#fff';
        youdaoConfigBtn.style.border = 'none';
        youdaoConfigBtn.style.borderRadius = '6px';
        youdaoConfigBtn.style.cursor = 'pointer';
        youdaoConfigBtn.onmouseenter = () => youdaoConfigBtn.style.background = '#c97c1a';
        youdaoConfigBtn.onmouseleave = () => youdaoConfigBtn.style.background = '#e67e22';
        youdaoConfigBtn.onclick = () => {
            new YoudaoTranslateConfigModal(this.app, this.plugin).open();
        };
        // 新增：复合翻译按钮
        const compositeBtn = containerEl.createEl('button', { text: '复合翻译设置' });
        compositeBtn.style.margin = '16px 0 16px 12px';
        compositeBtn.style.padding = '8px 24px';
        compositeBtn.style.fontSize = '1.1em';
        compositeBtn.style.background = '#0078d4';
        compositeBtn.style.color = '#fff';
        compositeBtn.style.border = 'none';
        compositeBtn.style.borderRadius = '6px';
        compositeBtn.style.cursor = 'pointer';
        compositeBtn.onmouseenter = () => compositeBtn.style.background = '#005a9e';
        compositeBtn.onmouseleave = () => compositeBtn.style.background = '#0078d4';
        compositeBtn.onclick = () => {
            new CompositeTranslateConfigModal(this.app, this.plugin).open();
        };
        // 插入到有道按钮右侧
        youdaoConfigBtn.parentElement?.insertBefore(compositeBtn, youdaoConfigBtn.nextSibling);

        // === 新增：翻译模式下拉框，合并自动切换逻辑 ===
        //let lastTranslateModel = this.plugin.settings.translateModel; // 放在 display() 作用域

        new Setting(containerEl)
            .setName("翻译模式")
            .setDesc("选择翻译模式")
            .addDropdown(drop => drop
                .addOption("youdao", "有道翻译")
                .addOption("composite", "复合翻译")
                .setValue(this.plugin.settings.translateModel || "youdao")
                .onChange(async (value) => {
                    // 先保存切换前的模式
                    const prevModel = lastTranslateModel;
                    lastTranslateModel = value; // 立即更新为新值
                    console.log('[调试] onChange触发，准备切换翻译模式，切换前模式:', prevModel, '切换后模式:', value);
                    this.plugin.settings.translateModel = value;
                    await this.plugin.saveSettings();
                    isSwitchingMode = true;
                    console.log('[调试] onChange: isSwitchingMode 设为 true');
                    refreshStatus();
                    try {
                        // 只要是从 composite 切到 youdao，先递增端口
                        if (prevModel === 'composite' && value === 'youdao') {
                            let basePort = parseInt(this.plugin.settings.serverPort) || 4000;
                            let newPort = await findAvailablePort(basePort + 1, 20);
                            console.log(`[切换模式] composite->youdao，分配新端口: ${newPort}`);
                            this.plugin.settings.serverPort = String(newPort);
                            await this.plugin.saveSettings();
                            refreshStatus('正在切换（分配新端口 ' + newPort + '）...');
                            await (this.plugin as any).stopServer(true);
                            await (this.plugin as any).startServer(true);
                            refreshStatus('新端口 ' + newPort + ' 已启动，继续切换...');
                        }
                        // 继续原有切换流程
                        await (this.plugin as any).stopServer(true);
                        await (this.plugin as any).startServer(true);
                    } finally {
                        isSwitchingMode = false;
                        console.log('[调试] onChange finally: isSwitchingMode 设为 false');
                        refreshStatus();
                    }
                })
            );

        const otherFuncBtn = containerEl.createEl('button', { text: '其他功能' });
        otherFuncBtn.style.background = '#fff';
        otherFuncBtn.style.color = '#000';
        otherFuncBtn.style.marginLeft = '12px';
        otherFuncBtn.onclick = () => {
            new OtherFunctionModal(this.app, this.plugin).open();
        };
        // 插入到复合翻译设置按钮右侧
        compositeBtn.parentElement?.insertBefore(otherFuncBtn, compositeBtn.nextSibling);
    }
}

// 1. 定义所有有道命令id
const YOUDAO_COMMAND_IDS = [
  'youdao-translate-selection-to-english',
  'youdao-translate-current-line'
];

// 2. 获取命令的快捷键组合
function getHotkeyCombosForCommand(app: App, commandId: string): string[][] {
    const hotkeyManager = (app as any).hotkeyManager;
    if (!hotkeyManager) return [];
    const custom = hotkeyManager.customKeys?.[commandId] || [];
    const builtIn = hotkeyManager.hotkeys?.[commandId] || [];
    const all = [...custom, ...builtIn];
    return all.map((h: any) => [...(h.modifiers || []), h.key]);
}

function matchHotkey(e: KeyboardEvent, combo: string[]): boolean {
    const ctrl = combo.includes('Ctrl');
    const shift = combo.includes('Shift');
    const alt = combo.includes('Alt');
    const mod = combo.includes('Mod');
    const key = combo[combo.length - 1];
    return (
        ctrl === e.ctrlKey &&
        shift === e.shiftKey &&
        alt === e.altKey &&
        (mod ? (e.metaKey || e.ctrlKey) : true) &&
        e.key.toUpperCase() === key.toUpperCase()
    );
}

// 3. 在弹窗 onOpen 时监听所有相关命令的快捷键
let globalMinimizedModals: { id: string, restoreBtn: HTMLElement, modal: Modal }[] = [];
function getUniqueModalId() {
    return 'modal-' + Math.random().toString(36).slice(2, 10) + '-' + Date.now();
}
class TranslateResultModal extends Modal {
    original: string;
    translated: string;
    color: string;
    mode: string;
    plugin: any;
    _keydownHandler: any;
    _isMinimized: boolean = false;
    _isMaximized: boolean = false;
    _modalId: string;
    _successCount?: number;
    _totalCount?: number;
    constructor(app: App, original: string, translated: string, color: string, mode: string, plugin: any, successCount?: number, totalCount?: number) {
        super(app);
        this.original = original;
        this.translated = translated;
        this.color = color;
        this.mode = mode;
        this.plugin = plugin;
        this._modalId = getUniqueModalId();
        this._successCount = successCount;
        this._totalCount = totalCount;
    }
    onOpen() {
        const { contentEl, modalEl } = this;
        contentEl.empty();
        contentEl.style.userSelect = "text";
        // 标题栏和按钮
        const header = contentEl.createEl("div");
        header.style.display = "flex";
        header.style.justifyContent = "space-between";
        header.style.alignItems = "center";
        header.style.marginBottom = "8px";
        const title = header.createEl("h2", { text: "翻译结果" });
        title.style.margin = "0";
        // 按钮组
        const btnGroup = header.createEl("div");
        btnGroup.style.display = "flex";
        btnGroup.style.gap = "12px";
        // --- 新增：翻译成功率 ---
        if (typeof this._successCount === 'number' && typeof this._totalCount === 'number' && this._totalCount > 0) {
            const rateDiv = header.createEl("div");
            rateDiv.textContent = `翻译成功率: ${this._successCount}/${this._totalCount}`;
            rateDiv.style.position = "absolute";
            rateDiv.style.right = "110px";
            rateDiv.style.top = "18px";
            rateDiv.style.fontWeight = "bold";
            rateDiv.style.color = "#e53935";
            rateDiv.style.fontSize = "1.08em";
            rateDiv.style.zIndex = "100000";
            // 保证不与按钮组重叠
        }
        // 最小化按钮
        const minBtn = btnGroup.createEl("button", { text: "–" });
        minBtn.title = "最小化";
        minBtn.style.fontSize = "1.3em";
        minBtn.style.width = "32px";
        minBtn.style.height = "32px";
        minBtn.style.border = "none";
        minBtn.style.background = "none";
        minBtn.style.cursor = "pointer";
        minBtn.onclick = () => {
            this._isMinimized = true;
            modalEl.style.display = "none";
            // 任务栏区域
            let bar = document.getElementById('youdao-modal-taskbar');
            if (!bar) {
                bar = document.createElement('div');
                bar.id = 'youdao-modal-taskbar';
                bar.style.position = 'fixed';
                bar.style.left = '50%';
                bar.style.transform = 'translateX(-50%)';
                bar.style.right = '';
                bar.style.bottom = '0';
                bar.style.height = '44px';
                bar.style.background = 'rgba(255,255,255,0.95)';
                bar.style.zIndex = '99999';
                bar.style.display = 'flex';
                bar.style.alignItems = 'center';
                bar.style.gap = '12px';
                bar.style.padding = '0 16px';
                document.body.appendChild(bar);
            }
            // 还原按钮
            const restoreBtn = document.createElement('button');
            restoreBtn.textContent = '翻译结果';
            restoreBtn.title = this.original.slice(0, 20) || '翻译结果';
            restoreBtn.style.margin = '0 8px';
            restoreBtn.style.padding = '6px 18px';
            restoreBtn.style.fontSize = '1em';
            restoreBtn.style.borderRadius = '6px';
            restoreBtn.style.background = '#f5f5f5';
            restoreBtn.style.border = '1px solid #ccc';
            restoreBtn.style.cursor = 'pointer';
            restoreBtn.onclick = () => {
                modalEl.style.display = '';
                this._isMinimized = false;
                restoreBtn.remove();
                // 移除全局记录
                globalMinimizedModals = globalMinimizedModals.filter(m => m.id !== this._modalId);
                // 如果任务栏无按钮则自动隐藏
                if (bar && bar.children.length === 0) bar.remove();
            };
            bar.appendChild(restoreBtn);
            globalMinimizedModals.push({ id: this._modalId, restoreBtn, modal: this });
        };
        // 最大化按钮
        const maxBtn = btnGroup.createEl("button", { text: "☐" });
        maxBtn.title = "最大化";
        maxBtn.style.fontSize = "1.1em";
        maxBtn.style.width = "32px";
        maxBtn.style.height = "32px";
        maxBtn.style.border = "none";
        maxBtn.style.background = "none";
        maxBtn.style.cursor = "pointer";
        maxBtn.onclick = () => {
            this._isMaximized = !this._isMaximized;
            if (this._isMaximized) {
                modalEl.style.width = '98vw';
                modalEl.style.height = '96vh';
                modalEl.style.left = '1vw';
                modalEl.style.top = '2vh';
                maxBtn.style.fontWeight = 'bold';
            } else {
                modalEl.style.width = '';
                modalEl.style.height = '';
                modalEl.style.left = '';
                modalEl.style.top = '';
                maxBtn.style.fontWeight = '';
            }
        };
        // 内容区
        const mainContent = contentEl.createEl("div");
        mainContent.className = 'modal-content';
        if (this.mode === "merge") {
            const originalLines = this.original.split(/\r?\n/);
            const translatedLines = this.translated.split(/\r?\n/);
            const maxLen = Math.max(originalLines.length, translatedLines.length);
            for (let i = 0; i < maxLen; i++) {
                if (originalLines[i]) {
                    const origDiv = mainContent.createEl("div", { text: originalLines[i] });
                    origDiv.style.margin = "4px 0";
                }
                if (translatedLines[i]) {
                    const transDiv = mainContent.createEl("div", { text: translatedLines[i] });
                    transDiv.style.color = this.color;
                    transDiv.style.margin = "0 0 8px 1em";
                }
            }
        } else {
            mainContent.createEl("div", { text: "原文：" });
            const origPre = mainContent.createEl("pre", { text: this.original });
            origPre.style.fontFamily = "inherit";
            origPre.style.fontSize = "1em";    // 可选，和交错显示一致

            mainContent.createEl("div", { text: "译文：" });
            const transPre = mainContent.createEl("pre", { text: this.translated });
            transPre.style.color = this.color;
            transPre.style.fontFamily = "inherit";
            transPre.style.fontSize = "1em";    // 可选，和交错显示一致
        }

        modalEl.style.resize = "both";
        modalEl.style.overflow = "auto";
        modalEl.style.position = "absolute";
        modalEl.style.background = "#fff";
        modalEl.style.zIndex = String(Date.now());
        makeModalDraggable(modalEl);
        // 动态获取快捷键，查不到用兜底
        let selectionHotkey = getFirstHotkeyCombo(this.app, 'youdao-translate-selection-to-english');
        let lineHotkey = getFirstHotkeyCombo(this.app, 'youdao-translate-current-line');
        // 新增：获取"始终机器翻译选中文本"命令的快捷键
        let forceMachineHotkey = getFirstHotkeyCombo(this.app, 'youdao-force-machine-translate-selection');
        this._keydownHandler = async (e: KeyboardEvent) => {
            console.log('[YoudaoPlugin] 弹窗收到 keydown:', e, 'selectionHotkey:', selectionHotkey, 'lineHotkey:', lineHotkey, 'forceMachineHotkey:', forceMachineHotkey);
            const sel = window.getSelection()?.toString().trim();
            // 检查"始终机器翻译选中文本"快捷键
            if (forceMachineHotkey &&
                e.ctrlKey === forceMachineHotkey.ctrl &&
                e.shiftKey === forceMachineHotkey.shift &&
                e.altKey === forceMachineHotkey.alt &&
                e.key.toLowerCase() === forceMachineHotkey.key
            ) {
                if (sel) {
                    // 复用主命令的批量机器翻译逻辑
                    const plugin = this.plugin;
                    const settings = plugin.settings;
                    const lines = sel.split(/\r?\n/).filter(line => line.trim().length > 0);
                    const getTargetLang = (text: string) => {
                        if (settings.isBiDirection) {
                            return /[\u4e00-\u9fa5]/.test(text)
                                ? (settings.translateModel === 'composite' ? 'en' : 'en')
                                : (settings.translateModel === 'composite'
                                    ? (settings.textTargetLang === 'zh' ? 'zh-Hans' : settings.textTargetLang)
                                    : (settings.textTargetLang === 'zh' ? 'zh-CHS' : settings.textTargetLang));
                        } else {
                            return settings.translateModel === 'composite'
                                ? (settings.textTargetLang === 'zh' ? 'zh-Hans' : settings.textTargetLang)
                                : (settings.textTargetLang === 'zh' ? 'zh-CHS' : settings.textTargetLang);
                        }
                    };
                    const color = settings.textTranslationColor;
                    const mode = settings.textTranslationMode;
                    if (settings.translateModel === 'composite') {
                        await unifiedBatchTranslateForMicrosoft({
                            app: this.app,
                            items: lines,
                            getTargetLang,
                            msAdapter: new MicrosoftAdapter(
                                settings.microsoftKey || '',
                                settings.microsoftRegion || '',
                                settings.microsoftEndpoint || '',
                                this.app
                            ),
                            color,
                            mode,
                            sleepInterval: settings.microsoftSleepInterval ?? 250,
                            plugin: plugin
                        });
                    } else {
                        await unifiedBatchTranslate({
                            app: this.app,
                            items: lines,
                            getTargetLang,
                            textAppKey: settings.textAppKey,
                            textAppSecret: settings.textAppSecret,
                            color,
                            mode,
                            port: settings.serverPort,
                            sleepInterval: settings.sleepInterval ?? 250,
                            plugin: plugin,
                            isBiDirection: settings.isBiDirection
                        });
                    }
                }
                e.preventDefault();
                return;
            }
            // 检查"选中内容"快捷键
            if (selectionHotkey && 
                e.ctrlKey === selectionHotkey.ctrl &&
                e.shiftKey === selectionHotkey.shift &&
                e.altKey === selectionHotkey.alt &&
                e.key.toLowerCase() === selectionHotkey.key
            ) {
                console.log('[YoudaoPlugin] 选中内容快捷键命中:', sel);
                if (sel) await handleDictOrTranslate(sel, this.app, this.plugin.settings, this.plugin);
                e.preventDefault();
                return;
            }
            // 检查"当前行"快捷键
            if (lineHotkey && 
                e.ctrlKey === lineHotkey.ctrl &&
                e.shiftKey === lineHotkey.shift &&
                e.altKey === lineHotkey.alt &&
                e.key.toLowerCase() === lineHotkey.key
            ) {
                console.log('[YoudaoPlugin] 当前行快捷键命中:', sel);
                if (sel) await handleDictOrTranslate(sel, this.app, this.plugin.settings, this.plugin);
                e.preventDefault();
                return;
            }
            console.log('[YoudaoPlugin] 未命中任何快捷键');
        };
        window.addEventListener('keydown', this._keydownHandler);
        // 强力兜底，禁止遮罩关闭
        const bg = this.modalEl.parentElement?.querySelector('.modal-bg') as HTMLElement;
        if (bg) {
            bg.style.pointerEvents = 'none';
        }
    }
    onClose() {
        window.removeEventListener('keydown', this._keydownHandler);
        this.contentEl.empty();
    }
    onClickOutside() {
        // 阻止点击遮罩关闭
    }
}

class AdvancedDictModal extends Modal {
    queryWord: string;
    headers: string[];
    rows: any[][];
    currentPage: number = 1;
    pageSize: number = 20;
    totalPages: number = 1;
    tableDiv: HTMLElement | null = null;
    _isMinimized: boolean = false;
    _isMaximized: boolean = false;
    _modalId: string;
    constructor(app: App, queryWord: string, headers: string[], rows: any[][]) {
        super(app);
        this.queryWord = queryWord;
        this.headers = headers;
        this.rows = rows;
        this.totalPages = Math.max(1, Math.ceil(rows.length / this.pageSize));
        this._modalId = getUniqueModalId();
    }
    onOpen() {
        const { contentEl, modalEl } = this;
        contentEl.empty();
        contentEl.style.userSelect = 'text';
        // 标题栏和按钮组
        const header = contentEl.createEl('div');
        header.style.display = 'flex';
        header.style.justifyContent = 'space-between';
        header.style.alignItems = 'center';
        header.style.marginBottom = '8px';
        const title = header.createEl('h2', { text: '高级释义' });
        title.style.margin = '0';
        // 按钮组
        const btnGroup = header.createEl('div');
        btnGroup.style.display = 'flex';
        btnGroup.style.gap = '12px';
        // 最小化按钮
        const minBtn = btnGroup.createEl('button', { text: '–' });
        minBtn.title = '最小化';
        minBtn.style.fontSize = '1.3em';
        minBtn.style.width = '32px';
        minBtn.style.height = '32px';
        minBtn.style.border = 'none';
        minBtn.style.background = 'none';
        minBtn.style.cursor = 'pointer';
        minBtn.onclick = () => {
            this._isMinimized = true;
            modalEl.style.display = 'none';
            let bar = document.getElementById('youdao-modal-taskbar');
            if (!bar) {
                bar = document.createElement('div');
                bar.id = 'youdao-modal-taskbar';
                bar.style.position = 'fixed';
                bar.style.left = '50%';
                bar.style.transform = 'translateX(-50%)';
                bar.style.right = '';
                bar.style.bottom = '0';
                bar.style.height = '44px';
                bar.style.background = 'rgba(255,255,255,0.95)';
                bar.style.zIndex = '99999';
                bar.style.display = 'flex';
                bar.style.alignItems = 'center';
                bar.style.gap = '12px';
                bar.style.padding = '0 16px';
                document.body.appendChild(bar);
            }
            const restoreBtn = document.createElement('button');
            restoreBtn.textContent = '高级释义';
            restoreBtn.title = '高级释义';
            restoreBtn.style.margin = '0 8px';
            restoreBtn.style.padding = '6px 18px';
            restoreBtn.style.fontSize = '1em';
            restoreBtn.style.borderRadius = '6px';
            restoreBtn.style.background = '#f5f5f5';
            restoreBtn.style.border = '1px solid #ccc';
            restoreBtn.style.cursor = 'pointer';
            restoreBtn.onclick = () => {
                modalEl.style.display = '';
                this._isMinimized = false;
                restoreBtn.remove();
                if (bar && bar.children.length === 0) bar.remove();
            };
            bar.appendChild(restoreBtn);
        };
        // 最大化按钮
        const maxBtn = btnGroup.createEl('button', { text: '☐' });
        maxBtn.title = '最大化';
        maxBtn.style.fontSize = '1.1em';
        maxBtn.style.width = '32px';
        maxBtn.style.height = '32px';
        maxBtn.style.border = 'none';
        maxBtn.style.background = 'none';
        maxBtn.style.cursor = 'pointer';
        maxBtn.onclick = () => {
            this._isMaximized = !this._isMaximized;
            if (this._isMaximized) {
                modalEl.style.width = '98vw';
                modalEl.style.height = '96vh';
                modalEl.style.left = '1vw';
                modalEl.style.top = '2vh';
                maxBtn.style.fontWeight = 'bold';
            } else {
                modalEl.style.width = '';
                modalEl.style.height = '';
                modalEl.style.left = '';
                modalEl.style.top = '';
                maxBtn.style.fontWeight = '';
            }
        };
        // 内容区
        if (!this.rows || this.rows.length === 0) {
            contentEl.createEl('div', { text: '无法在dict_service路径下的en-zh.sqlite3数据库中的wordW表的word列找到该单词或短语' });
            return;
        }
        // 表格区
        this.tableDiv = contentEl.createDiv();
        this.renderTable();
        this.renderPagination();
        // 只允许按住标题文字拖拽
        makeModalDraggable(modalEl, title);
    }
    renderTable() {
        if (!this.tableDiv) return;
        this.tableDiv.empty();
        // 复用renderHtmlTable，所有列宽/行高可拖拽
        this.tableDiv.appendChild(renderHtmlTable(
            this.headers,
            this.rows.slice((this.currentPage - 1) * this.pageSize, Math.min(this.currentPage * this.pageSize, this.rows.length)),
            [],
            [],
            () => {},
            () => {},
            false,
            (this.currentPage - 1) * this.pageSize
        ));
    }
    renderPagination() {
        if (!this.tableDiv) return;
        // 移除旧分页
        const oldBar = this.tableDiv.querySelector('.pagination-bar');
        if (oldBar) oldBar.remove();
        // 分页控件
        const bar = document.createElement('div');
        bar.className = 'pagination-bar';
        bar.style.display = 'flex';
        bar.style.justifyContent = 'flex-start';
        bar.style.alignItems = 'center';
        bar.style.margin = '12px 0 0 0';
        bar.style.width = '100%';
        // 页码信息
        const infoSpan = document.createElement('span');
        infoSpan.textContent = `第 ${this.currentPage} / ${this.totalPages} 页`;
        infoSpan.style.marginRight = '12px';
        bar.appendChild(infoSpan);
        // 页码输入框
        const pageInput = document.createElement('input');
        pageInput.type = 'number';
        pageInput.min = '1';
        pageInput.max = String(this.totalPages);
        pageInput.value = String(this.currentPage);
        pageInput.style.width = '48px';
        pageInput.style.marginRight = '12px';
        pageInput.style.fontSize = '1em';
        pageInput.style.verticalAlign = 'middle';
        pageInput.title = '跳转到指定页';
        pageInput.onkeydown = (e) => {
            if (e.key === 'Enter') {
                let val = parseInt(pageInput.value);
                if (isNaN(val) || val < 1) val = 1;
                if (val > this.totalPages) val = this.totalPages;
                if (val !== this.currentPage) {
                    this.currentPage = val;
                    this.renderTable();
                    this.renderPagination();
                }
            }
        };
        pageInput.onblur = () => {
            let val = parseInt(pageInput.value);
            if (isNaN(val) || val < 1) val = 1;
            if (val > this.totalPages) val = this.totalPages;
            if (val !== this.currentPage) {
                this.currentPage = val;
                this.renderTable();
                this.renderPagination();
            }
        };
        bar.appendChild(pageInput);
        // 上一页
        const prevBtn = document.createElement('button');
        prevBtn.textContent = '上一页';
        prevBtn.disabled = this.currentPage === 1;
        prevBtn.onclick = () => {
            if (this.currentPage > 1) {
                this.currentPage--;
                this.renderTable();
                this.renderPagination();
            }
        };
        bar.appendChild(prevBtn);
        // 下一页
        const nextBtn = document.createElement('button');
        nextBtn.textContent = '下一页';
        nextBtn.disabled = this.currentPage === this.totalPages;
        nextBtn.onclick = () => {
            if (this.currentPage < this.totalPages) {
                this.currentPage++;
                this.renderTable();
                this.renderPagination();
            }
        };
        bar.appendChild(nextBtn);
        this.tableDiv.appendChild(bar);
    }
}

class DictResultModal extends Modal {
    html: string;
    plugin: any;
    _keydownHandler: any;
    _isMinimized: boolean = false;
    _isMaximized: boolean = false;
    _modalId: string;
    constructor(app: App, html: string, plugin: any) {
        super(app);
        this.html = html;
        this.plugin = plugin;
        this._modalId = getUniqueModalId();
    }
    onOpen() {
        const { contentEl, modalEl } = this;
        contentEl.empty();
        contentEl.style.userSelect = "text";
        // 标题栏
        const header = contentEl.createEl("div");
        header.style.display = "flex";
        header.style.justifyContent = "center";
        header.style.alignItems = "center";
        header.style.position = "relative";
        header.style.marginBottom = "8px";
        const title = header.createEl("h2", { text: "词典释义" });
        title.style.margin = "0";
        // 按钮组（插入到 modalEl，绝对定位右上角，避开关闭按钮）
        const btnGroup = document.createElement("div");
        btnGroup.style.display = "flex";
        btnGroup.style.gap = "12px";
        btnGroup.style.position = "absolute";
        btnGroup.style.top = "12px";
        btnGroup.style.right = "48px"; // 留出关闭按钮空间
        // 最小化按钮
        const minBtn = document.createElement("button");
        minBtn.textContent = "–";
        minBtn.title = "最小化";
        minBtn.style.fontSize = "1.3em";
        minBtn.style.width = "32px";
        minBtn.style.height = "32px";
        minBtn.style.border = "none";
        minBtn.style.background = "none";
        minBtn.style.cursor = "pointer";
        minBtn.onclick = () => {
            this._isMinimized = true;
            modalEl.style.display = "none";
            let bar = document.getElementById('youdao-modal-taskbar');
            if (!bar) {
                bar = document.createElement('div');
                bar.id = 'youdao-modal-taskbar';
                bar.style.position = 'fixed';
                bar.style.left = '50%';
                bar.style.transform = 'translateX(-50%)';
                bar.style.right = '';
                bar.style.bottom = '0';
                bar.style.height = '44px';
                bar.style.background = 'rgba(255,255,255,0.95)';
                bar.style.zIndex = '99999';
                bar.style.display = 'flex';
                bar.style.alignItems = 'center';
                bar.style.gap = '12px';
                bar.style.padding = '0 16px';
                document.body.appendChild(bar);
            }
            const restoreBtn = document.createElement('button');
            restoreBtn.textContent = '词典释义';
            restoreBtn.title = '词典释义';
            restoreBtn.style.margin = '0 8px';
            restoreBtn.style.padding = '6px 18px';
            restoreBtn.style.fontSize = '1em';
            restoreBtn.style.borderRadius = '6px';
            restoreBtn.style.background = '#f5f5f5';
            restoreBtn.style.border = '1px solid #ccc';
            restoreBtn.style.cursor = 'pointer';
            restoreBtn.onclick = () => {
                modalEl.style.display = '';
                this._isMinimized = false;
                restoreBtn.remove();
                globalMinimizedModals = globalMinimizedModals.filter(m => m.id !== this._modalId);
                if (bar && bar.children.length === 0) bar.remove();
            };
            bar.appendChild(restoreBtn);
            globalMinimizedModals.push({ id: this._modalId, restoreBtn, modal: this });
        };
        btnGroup.appendChild(minBtn);
        // 最大化按钮
        const maxBtn = document.createElement("button");
        maxBtn.textContent = "☐";
        maxBtn.title = "最大化";
        maxBtn.style.fontSize = "1.1em";
        maxBtn.style.width = "32px";
        maxBtn.style.height = "32px";
        maxBtn.style.border = "none";
        maxBtn.style.background = "none";
        maxBtn.style.cursor = "pointer";
        maxBtn.onclick = () => {
            this._isMaximized = !this._isMaximized;
            if (this._isMaximized) {
                modalEl.style.width = '98vw';
                modalEl.style.height = '96vh';
                modalEl.style.left = '1vw';
                modalEl.style.top = '2vh';
                maxBtn.style.fontWeight = 'bold';
            } else {
                modalEl.style.width = '';
                modalEl.style.height = '';
                modalEl.style.left = '';
                modalEl.style.top = '';
                maxBtn.style.fontWeight = '';
            }
        };
        btnGroup.appendChild(maxBtn);
        modalEl.appendChild(btnGroup); // 关键：插入到 modalEl
        // 新增"高级"按钮，插入在标题栏下方、内容上方
        const advancedBtn = contentEl.createEl('button', {
            text: '高级',
            cls: 'dict-advanced-btn'
        });
        advancedBtn.setAttr('style', 'margin: 8px 0 12px 0; background: #fff; color: #000; border: 1px solid #ccc; border-radius: 4px; padding: 4px 16px; cursor: pointer;');
        advancedBtn.onclick = async () => {
            // 获取当前查词内容
            let queryWord = '';
            // 尝试从html中提取
            const m = this.html.match(/<div[^>]*font-size:1.5em[^>]*margin:8px 0 8px 0[^>]*>(.*?)<\/div>/);
            if (m && m[1]) queryWord = m[1].trim();
            if (!queryWord) {
                new Notice('无法获取当前查词内容');
            return;
        }
            // 读取sqlite3文件
            let dbFilePath = '';
            const dictServicePath = this.plugin.settings.dictServicePath;
            if (!dictServicePath) {
                new Notice('请先在设置中填写dict_service路径');
                return;
            }
            dbFilePath = dictServicePath.replace(/[/\\]+$/, '') + '/en-zh.sqlite3';
            let arrayBuffer;
            try {
                arrayBuffer = fs.readFileSync(dbFilePath);
            } catch (e) {
                new Notice('无法读取数据库: ' + dbFilePath);
                return;
            }
            // 加载sql.js
            const initSqlJs = await loadSqlJs();
            const SQL = await initSqlJs({ locateFile: (file: string) => 'https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/' + file });
            const db = new SQL.Database(new Uint8Array(arrayBuffer));
            // 查询wordW表
            let rows = [];
            let headers = ['word', 'Interpretation_serial_number', 'Part_of_speech', 'Synonyms', 'Interpretation', 'Example_sentence', 'Related_phrases_derivatives'];
            try {
                const res = db.exec(`SELECT * FROM wordW WHERE word = ?`, [queryWord]);
                if (res && res[0]) {
                    headers = res[0].columns;
                    rows = res[0].values;
                }
            } catch (e) {
                new Notice('查询wordW表失败');
                return;
            }
            // 弹窗展示
            new AdvancedDictModal(this.plugin.app, queryWord, headers, rows).open();
        };
        // 内容区
        const mainContent = contentEl.createEl("div");
        mainContent.className = 'modal-content';
        mainContent.innerHTML = this.html;
        modalEl.style.resize = "both";
        modalEl.style.overflow = "auto";
        modalEl.style.position = "absolute";
        modalEl.style.background = "#fff";
        modalEl.style.zIndex = String(Date.now());
        // 只允许按住标题拖拽
        makeModalDraggable(modalEl, title);
        // 动态获取快捷键，查不到用兜底
        let selectionHotkey = getFirstHotkeyCombo(this.app, 'youdao-translate-selection-to-english');
        let lineHotkey = getFirstHotkeyCombo(this.app, 'youdao-translate-current-line');
        // 新增：获取"始终机器翻译选中文本"命令的快捷键
        let forceMachineHotkey = getFirstHotkeyCombo(this.app, 'youdao-force-machine-translate-selection');
        this._keydownHandler = async (e: KeyboardEvent) => {
            console.log('[YoudaoPlugin] 弹窗收到 keydown:', e, 'selectionHotkey:', selectionHotkey, 'lineHotkey:', lineHotkey, 'forceMachineHotkey:', forceMachineHotkey);
            const sel = window.getSelection()?.toString().trim();
            // 检查"始终机器翻译选中文本"快捷键
            if (forceMachineHotkey && 
                e.ctrlKey === forceMachineHotkey.ctrl &&
                e.shiftKey === forceMachineHotkey.shift &&
                e.altKey === forceMachineHotkey.alt &&
                e.key.toLowerCase() === forceMachineHotkey.key
            ) {
                if (sel) {
                    // 复用主命令的批量机器翻译逻辑
                    const plugin = this.plugin;
                    const settings = plugin.settings;
                    const lines = sel.split(/\r?\n/).filter(line => line.trim().length > 0);
                    const getTargetLang = (text: string) => {
                        if (settings.isBiDirection) {
                            return /[\u4e00-\u9fa5]/.test(text)
                                ? (settings.translateModel === 'composite' ? 'en' : 'en')
                                : (settings.translateModel === 'composite'
                                    ? (settings.textTargetLang === 'zh' ? 'zh-Hans' : settings.textTargetLang)
                                    : (settings.textTargetLang === 'zh' ? 'zh-CHS' : settings.textTargetLang));
                        } else {
                            return settings.translateModel === 'composite'
                                ? (settings.textTargetLang === 'zh' ? 'zh-Hans' : settings.textTargetLang)
                                : (settings.textTargetLang === 'zh' ? 'zh-CHS' : settings.textTargetLang);
                        }
                    };
                    const color = settings.textTranslationColor;
                    const mode = settings.textTranslationMode;
                    if (settings.translateModel === 'composite') {
                        await unifiedBatchTranslateForMicrosoft({
                            app: this.app,
                            items: lines,
                            getTargetLang,
                            msAdapter: new MicrosoftAdapter(
                                settings.microsoftKey || '',
                                settings.microsoftRegion || '',
                                settings.microsoftEndpoint || '',
                                this.app
                            ),
                            color,
                            mode,
                            sleepInterval: settings.microsoftSleepInterval ?? 250,
                            plugin: plugin
                        });
                    } else {
                        await unifiedBatchTranslate({
                            app: this.app,
                            items: lines,
                            getTargetLang,
                            textAppKey: settings.textAppKey,
                            textAppSecret: settings.textAppSecret,
                            color,
                            mode,
                            port: settings.serverPort,
                            sleepInterval: settings.sleepInterval ?? 250,
                            plugin: plugin,
                            isBiDirection: settings.isBiDirection
                        });
                    }
                }
                e.preventDefault();
                return;
            }
            // 检查"选中内容"快捷键
            if (selectionHotkey && 
                e.ctrlKey === selectionHotkey.ctrl &&
                e.shiftKey === selectionHotkey.shift &&
                e.altKey === selectionHotkey.alt &&
                e.key.toLowerCase() === selectionHotkey.key
            ) {
                console.log('[YoudaoPlugin] 选中内容快捷键命中:', sel);
                if (sel) await handleDictOrTranslate(sel, this.app, this.plugin.settings, this.plugin);
                e.preventDefault();
                return;
            }
            // 检查"当前行"快捷键
            if (lineHotkey && 
                e.ctrlKey === lineHotkey.ctrl &&
                e.shiftKey === lineHotkey.shift &&
                e.altKey === lineHotkey.alt &&
                e.key.toLowerCase() === lineHotkey.key
            ) {
                console.log('[YoudaoPlugin] 当前行快捷键命中:', sel);
                if (sel) await handleDictOrTranslate(sel, this.app, this.plugin.settings, this.plugin);
                e.preventDefault();
                return;
            }
            console.log('[YoudaoPlugin] 未命中任何快捷键');
        };
        window.addEventListener('keydown', this._keydownHandler);
        // 强力兜底，禁止遮罩关闭
        const bg = this.modalEl.parentElement?.querySelector('.modal-bg') as HTMLElement;
        if (bg) {
            bg.style.pointerEvents = 'none';
        }
    }
    onClose() {
        window.removeEventListener('keydown', this._keydownHandler);
        this.contentEl.empty();
    }
    onClickOutside() {
        // 阻止点击遮罩关闭
    }
}

function xhrPost(url: string, params: string): Promise<any> {
    return new Promise((resolve, reject) => {
        const xhr = new XMLHttpRequest();
        xhr.open("POST", url);
        xhr.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
        xhr.onreadystatechange = function () {
            if (xhr.readyState === 4) {
                if (xhr.status === 200) {
                    try {
                        resolve(JSON.parse(xhr.responseText));
                    } catch (e) {
                        reject(e);
                    }
                } else {
                    reject(xhr.statusText);
                }
            }
        };
        xhr.onerror = reject;
        xhr.send(params);
    });
}

async function youdaoTranslate(
    q: string,
    appKey: string,
    appSecret: string,
    to: string,
    app: App,
    port: string
): Promise<string | null> {
    if (!appKey || !appSecret) return null;
    const salt = Date.now().toString();
    const curtime = Math.round(Date.now() / 1000).toString();
    const str1 = appKey + truncate(q) + salt + curtime + appSecret;
    const sign = CryptoJS.SHA256(str1).toString(CryptoJS.enc.Hex);

    const params = new URLSearchParams({
        q,
        appKey,
        salt,
        from: "auto",
        to,
        sign,
        signType: "v3",
        curtime
    }).toString();

    try {
        const data = await xhrPost(`http://127.0.0.1:${port}/youdao`, params);
        if (data && data.translation && data.translation[0]) {
            return data.translation[0];
        }
        return null;
    } catch (e) {
        return null;
    }
}

function truncate(q: string): string {
    const len = q.length;
    if (len <= 20) return (q ?? '').toString();
    const str = (q ?? '').toString();
    return str.substring(0, 10) + len + str.substring(str.length - 10, str.length);
}

// 1. 智能主语言检测
function detectMainLangSmart(text: string): 'zh' | 'en' {
    // 去除标点符号
    const clean = text.replace(/[\p{P}\p{S}]/gu, '');
    const enWords = (clean.match(/[a-zA-Z]+/g) || []).length;
    const zhChars = (clean.match(/[\u4e00-\u9fa5]/g) || []).length;
    return enWords >= zhChars ? 'en' : 'zh';
}

// 2. OCR分段合并与过滤
function mergeAndFilterOcrRegions(regions: string[]): string[] {
    const merged: string[] = [];
    let buffer = '';
    for (let seg of regions) {
        seg = (seg ?? '').toString().trim();
        if (!seg) continue;
        // 过滤全是标点或空白
        if (/^[\\s\\p{P}]+$/u.test(seg)) continue;
        // 合并过短的段
        if (seg.length < 5 && merged.length > 0) {
            merged[merged.length - 1] += ' ' + seg;
        } else {
            merged.push(seg);
        }
    }
    return merged;
}

// 3. 优化后的图片翻译主逻辑
async function translateImageFile(
    app: App,
    file: TFile,
    imageAppKey: string,
    imageAppSecret: string,
    textAppKey: string,
    textAppSecret: string,
    to: string,
    color: string,
    mode: string,
    port: string,
    plugin?: any // 新增参数，传递插件实例
) {
    if (hasMinimizedModal()) {
        new Notice("请先还原/关闭顶层弹窗，再进行其他操作");
        return;
    }
    // loading动画提前，且样式全局只插入一次
    let loading = document.createElement('div');
    loading.className = 'youdao-loading';
    loading.innerHTML = `<div class='youdao-spinner'></div> 翻译正在进行，请稍等...<div class='youdao-progress' style='margin-top:16px;font-size:1.1em;'></div>`;
    Object.assign(loading.style, {
      position: 'fixed', left: '50%', top: '30%', transform: 'translate(-50%, -50%)',
      zIndex: 99999, background: '#fff', padding: '32px 48px', borderRadius: '12px',
      boxShadow: '0 2px 16px #0002', fontSize: '1.2em', textAlign: 'center'
    });
    document.body.appendChild(loading);
    const progressEl = loading.querySelector('.youdao-progress');
    if (!document.getElementById('youdao-spinner-style')) {
        const style = document.createElement('style');
        style.id = 'youdao-spinner-style';
        style.innerHTML = `.youdao-spinner{display:inline-block;width:32px;height:32px;border:4px solid #eee;border-top:4px solid #1a73e8;border-radius:50%;animation:spin 1s linear infinite;vertical-align:middle;margin-right:16px;}@keyframes spin{0%{transform:rotate(0)}100%{transform:rotate(360deg)}}`;
        document.head.appendChild(style);
    }
    try {
        const isBiDirection = (plugin && plugin.settings && typeof plugin.settings.isBiDirection === 'boolean')
            ? plugin.settings.isBiDirection
            : (app as any).plugins?.plugins?.youdaoTranslatePlugin?.settings?.isBiDirection;
    const arrayBuffer = await app.vault.readBinary(file);
    const q = arrayBufferToBase64(arrayBuffer);
        // OCR识别所有段落内容（始终用图片翻译key/secret）
        const formData = new URLSearchParams();
        formData.append("type", "1");
        formData.append("from", "auto");
        formData.append("to", "en");
        formData.append("appKey", imageAppKey); // 用图片key
    const salt = Date.now().toString();
    const curtime = Math.floor(Date.now() / 1000).toString();
    let input = q;
    if ((q ?? '').toString().length > 20) {
        const str = (q ?? '').toString();
        input = str.substring(0, 10) + str.length + str.substring(str.length - 10, str.length);
    }
        const signStr = imageAppKey + input + salt + curtime + imageAppSecret; // 用图片secret
    const sign = CryptoJS.SHA256(signStr).toString(CryptoJS.enc.Hex);
    formData.append("salt", salt);
    formData.append("sign", sign);
    formData.append("signType", "v3");
    formData.append("curtime", curtime);
    formData.append("q", q);
        let ocrRegions = [];
    try {
        const resp = await fetch(`http://127.0.0.1:${port}/youdao-image`, {
            method: "POST",
            headers: { "Content-Type": "application/x-www-form-urlencoded" },
            body: formData.toString()
        });
        const data = await resp.json();
        if (data && data.resRegions) {
                ocrRegions = data.resRegions.map((r: any) => r.context);
        }
    } catch (e) {
            new Notice("图片内容识别失败: " + e);
            return;
        }
        if (!ocrRegions.length) {
            new Notice("未识别到图片内容");
            return;
        }
        ocrRegions = mergeAndFilterOcrRegions(ocrRegions);
        // --- 统一：无论 isBiDirection 开关状态，均分块分别翻译 ---
        await unifiedBatchTranslate({
            app,
            items: ocrRegions,
            getTargetLang: (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-CHS',
            textAppKey,
            textAppSecret,
            color,
            mode,
            port,
            sleepInterval: plugin && plugin.settings ? plugin.settings.sleepInterval : 250,
            plugin,
            isBiDirection
        });
    } finally {
        if (loading) loading.remove();
    }
}

function arrayBufferToBase64(buffer: ArrayBuffer): string {
    let binary = '';
    const bytes = new Uint8Array(buffer);
    const len = bytes.byteLength;
    for (let i = 0; i < len; i++) {
        binary += String.fromCharCode(bytes[i]);
    }
    return btoa(binary);
}

function makeModalDraggable(modalEl: HTMLElement, dragHandle?: HTMLElement) {
    let isDragging = false;
    let startX = 0, startY = 0, startLeft = 0, startTop = 0;
    modalEl.style.pointerEvents = 'auto';
    if (!(window as any).__modalZIndex) (window as any).__modalZIndex = 100000;
    (window as any).__modalZIndex++;
    modalEl.style.zIndex = String((window as any).__modalZIndex);
    const bg = modalEl.parentElement?.querySelector('.modal-bg') as HTMLElement;
    if (bg) bg.style.display = 'none';
    // 支持 dragHandle
    let handle = dragHandle || modalEl.querySelector("h2") || modalEl;
    handle.style.cursor = "move";
    handle.addEventListener("mousedown", (e) => {
        isDragging = true;
        startX = e.clientX;
        startY = e.clientY;
        // 修复：拖动前强制 position: absolute 并赋 left/top，移除 margin/transform
        const rect = modalEl.getBoundingClientRect();
        modalEl.style.position = "absolute";
        modalEl.style.left = rect.left + "px";
        modalEl.style.top = rect.top + "px";
        modalEl.style.margin = "0";
        modalEl.style.transform = "none";
        startLeft = rect.left;
        startTop = rect.top;
        document.body.style.userSelect = "none";
        (window as any).__modalZIndex++;
        modalEl.style.zIndex = String((window as any).__modalZIndex);
    });
    window.addEventListener("mousemove", (e) => {
        if (!isDragging) return;
        const dx = e.clientX - startX;
        const dy = e.clientY - startY;
        modalEl.style.left = `${startLeft + dx}px`;
        modalEl.style.top = `${startTop + dy}px`;
    });
    window.addEventListener("mouseup", () => {
        isDragging = false;
        document.body.style.userSelect = "";
    });
}

class HowToUseModal extends Modal {
    _isMinimized: boolean = false;
    _isMaximized: boolean = false;
    _modalId: string;
    constructor(app: App) {
        super(app);
        this._modalId = getUniqueModalId();
    }
    onOpen() {
        const { contentEl, modalEl } = this;
        contentEl.empty();
        contentEl.style.userSelect = "text";
        contentEl.createEl("h2", { text: "插件使用说明" });
        contentEl.createEl("div", {
            text: `用鼠标长按翻译结果可以实现拖拽效果 \n\n- 支持文本和图片翻译\n- 设置目标语言和译文颜色\n- 选中内容后使用命令或快捷键翻译\n- ...`
        });
        modalEl.style.width = "800px";
        modalEl.style.height = "600px";
        modalEl.style.maxWidth = "90vw";
        modalEl.style.maxHeight = "90vh";
        modalEl.style.overflow = "auto";
        makeModalDraggable(modalEl);
        // 最小化/最大化按钮
        const header = contentEl.querySelector('h2') || contentEl.firstElementChild;
        if (header) {
            const btnGroup = document.createElement('div');
            btnGroup.style.display = 'flex';
            btnGroup.style.gap = '8px';
            btnGroup.style.position = 'absolute';
            btnGroup.style.top = '16px';
            btnGroup.style.right = '32px';
            // 最小化
            const minBtn = document.createElement('button');
            minBtn.textContent = '–';
            minBtn.title = '最小化';
            minBtn.onclick = () => {
                this._isMinimized = true;
                modalEl.style.display = 'none';
                let bar = document.getElementById('youdao-modal-taskbar');
                if (!bar) {
                    bar = document.createElement('div');
                    bar.id = 'youdao-modal-taskbar';
                    bar.style.position = 'fixed';
                    bar.style.left = '50%';
                    bar.style.transform = 'translateX(-50%)';
                    bar.style.bottom = '0';
                    bar.style.height = '44px';
                    bar.style.background = 'rgba(255,255,255,0.95)';
                    bar.style.zIndex = '99999';
                    bar.style.display = 'flex';
                    bar.style.alignItems = 'center';
                    bar.style.gap = '12px';
                    bar.style.padding = '0 16px';
                    document.body.appendChild(bar);
                }
                const restoreBtn = document.createElement('button');
                restoreBtn.textContent = '使用说明';
                restoreBtn.onclick = () => {
                    modalEl.style.display = '';
                    this._isMinimized = false;
                    restoreBtn.remove();
                };
                bar.appendChild(restoreBtn);
            };
            // 最大化
            const maxBtn = document.createElement('button');
            maxBtn.textContent = '☐';
            maxBtn.title = '最大化';
            maxBtn.onclick = () => {
                this._isMaximized = !this._isMaximized;
                if (this._isMaximized) {
                    modalEl.style.width = '98vw';
                    modalEl.style.height = '96vh';
                    modalEl.style.left = '1vw';
                    modalEl.style.top = '2vh';
                    maxBtn.style.fontWeight = 'bold';
                } else {
                    modalEl.style.width = '';
                    modalEl.style.height = '';
                    modalEl.style.left = '';
                    modalEl.style.top = '';
                    maxBtn.style.fontWeight = '';
                }
            };
            btnGroup.appendChild(minBtn);
            btnGroup.appendChild(maxBtn);
            contentEl.appendChild(btnGroup);
        }
        // 禁止遮罩关闭
        const bg = modalEl.parentElement?.querySelector('.modal-bg') as HTMLElement;
        if (bg) bg.style.pointerEvents = 'none';
        // === 拖动标志插入 start ===
        let dragHandle = modalEl.querySelector('.modal-drag-handle') as HTMLElement;
        if (!dragHandle) {
            dragHandle = document.createElement('div');
            dragHandle.className = 'modal-drag-handle';
            dragHandle.textContent = '≡';
            dragHandle.style.position = 'absolute';
            dragHandle.style.left = '16px';
            dragHandle.style.top = '16px';
            dragHandle.style.fontSize = '1.5em';
            dragHandle.style.cursor = 'grab';
            dragHandle.style.userSelect = 'none';
            dragHandle.style.zIndex = '100001';
            modalEl.appendChild(dragHandle);
        }
        makeModalDraggable(modalEl, dragHandle);
    }
    onClose() {
        this.contentEl.empty();
    }
}

// 新增：有道词典网页版API请求
async function youdaoDictLookup(q: string, port: string): Promise<any> {
    const resp = await fetch(`http://127.0.0.1:${port}/youdao-dict?q=${encodeURIComponent(q)}`);
    if (!resp.ok) throw new Error('词典API请求失败');
    return await resp.json();
}

function renderDictResult(dictData: any, lang: 'zh' | 'en', color: string, port: string, machineTranslation?: string, originalText?: string): string {
    let html = '';
    html += `<div style="font-size:1.5em;font-weight:bold;">词典释义</div>`;

    // ====== 新增：本地 definitions 兼容有道结构 ======
    // 如果是复合翻译（本地词典），将 definitions 映射为 ec.word[0].trs
    let w = null;
    let isLocalDict = false;
    if (
        dictData.definitions && Array.isArray(dictData.definitions) && dictData.definitions.length > 0 &&
        !(dictData.ec && dictData.ec.word && dictData.ec.word.length > 0)
    ) {
        isLocalDict = true;
        // 构造有道风格的 ec.word[0]
        w = {
            trs: dictData.definitions.map((item: any) => ({
                tr: [{
                    l: {
                        i: [
                            item.en
                                ? item.en
                                : (item.zh
                                    ? item.zh
                                    : (item.pos ? `[${item.pos}] ` : '') + (item.sense || '') + (item.trans_list ? ` ${item.trans_list}` : '')
                                )
                        ]
                    }
                }]
            })),
            // 例句
            exam_sents: Array.isArray(dictData.examples) ? dictData.examples.map((ex: any) => ({ eng: ex.en, chn: ex.zh })) : [],
        };
        // 网络释义
        dictData.web_trans = dictData.web_trans || {};
        if (Array.isArray(dictData.phrases) && dictData.phrases.length > 0) {
            dictData.web_trans['web-translation'] = dictData.phrases.map((item: any) => ({
                key: typeof item === 'string' ? item : (item.key || ''),
                trans: typeof item === 'string' ? [] : (item.trans ? [{ value: item.trans }] : [])
            }));
        }
    } else {
        // 原有有道风格
        if (dictData.ec && dictData.ec.word && dictData.ec.word.length > 0) {
            w = dictData.ec.word[0];
        } else if (dictData.ce && dictData.ce.word && dictData.ce.word.length > 0) {
            w = dictData.ce.word[0];
        } else if (dictData.ee && dictData.ee.word && dictData.ee.word.length > 0) {
            w = dictData.ee.word[0];
        }
    }
    // ====== END ======

    // 优先显示原文（加粗变大），无则用词典返回的单词
    let word = '';
    if (w) {
        if (typeof w['return-phrase'] === 'string') {
            word = w['return-phrase'];
        } else if (Array.isArray(w['return-phrase'])) {
            word = w['return-phrase'].join(' ');
        } else if (w['return-phrase']?.l?.i?.[0]) {
            word = w['return-phrase'].l.i[0];
        } else if (w['return-phrase']?.l?.i) {
            word = w['return-phrase'].l.i;
        } else if (w['return-phrase']) {
            word = String(w['return-phrase']);
        }
        if (word.length <= 1 && dictData.input) {
            word = dictData.input;
        }
    }
    const displayWord = originalText || word || '';
    html += `<div style="font-size:1.5em;font-weight:bold;margin:8px 0 8px 0">${displayWord}</div>`;

    // 音标和音频（仅有道词典）
    if (w && !isLocalDict) {
        const ukphone = w.ukphone || '';
        const usphone = w.usphone || '';
        const ukspeech = w.ukspeech || '';
        const usspeech = w.usspeech || '';
        const ukspeechUrl = ukspeech ? `http://127.0.0.1:${port}/youdao-audio?url=${getFullAudioUrl(ukspeech)}` : '';
        const usspeechUrl = usspeech ? `http://127.0.0.1:${port}/youdao-audio?url=${getFullAudioUrl(usspeech)}` : '';
        if (ukphone || usphone) {
            html += `<div>英[${ukphone}] ${ukspeechUrl ? `<audio src="${ukspeechUrl}" controls style="height:1em;vertical-align:middle"></audio>` : ''} 美[${usphone}] ${usspeechUrl ? `<audio src="${usspeechUrl}" controls style="height:1em;vertical-align:middle"></audio>` : ''}</div>`;
        }
    }

    // 基本释义
    let trsList = [];
    if (w && w.trs && w.trs.length > 0) {
        trsList = w.trs;
    } else {
        if (dictData.ec && dictData.ec.word && dictData.ec.word[0] && dictData.ec.word[0].trs && dictData.ec.word[0].trs.length > 0) {
            trsList = dictData.ec.word[0].trs;
        } else if (dictData.ce && dictData.ce.word && dictData.ce.word[0] && dictData.ce.word[0].trs && dictData.ce.word[0].trs.length > 0) {
            trsList = dictData.ce.word[0].trs;
        } else if (dictData.ee && dictData.ee.word && dictData.ee.word[0] && dictData.ee.word[0].trs && dictData.ee.word[0].trs.length > 0) {
            trsList = dictData.ee.word[0].trs;
        }
    }
    function isValidTrString(str: string): boolean {
        if (!str) return false;
        const s = str.trim().toLowerCase();
        if (!s) return false;
        if (s === 'link') return false;
        if (s.startsWith('app:ds:')) return false;
        if (/^[.:;\-]+$/.test(s)) return false;
        return true;
    }
    if (trsList.length > 0) {
        html += `<div style=\"margin-top:8px;\"><b>基本释义：</b><ul style=\"margin:8px 0 8px 1.5em;padding:0;\">`;
        trsList.forEach((tr: any) => {
            if (tr.tr && tr.tr[0] && tr.tr[0].l && tr.tr[0].l.i) {
                const item = tr.tr[0].l.i;
                function extractStrings(obj: any): string[] {
                    let result: string[] = [];
                    if (typeof obj === 'string') {
                        result.push(obj);
                    } else if (Array.isArray(obj)) {
                        obj.forEach(sub => {
                            result = result.concat(extractStrings(sub));
                        });
                    } else if (typeof obj === 'object' && obj !== null) {
                        for (const key in obj) {
                            result = result.concat(extractStrings(obj[key]));
                        }
                    }
                    return result;
                }
                extractStrings(item).forEach(str => {
                    if (isValidTrString(str)) {
                        html += `<li style=\"margin:6px 0;font-size:1.08em;line-height:1.7;\">${str}</li>`;
                    }
                });
            }
        });
        html += `</ul></div>`;
    } else {
        html += `<div style='margin:12px 0;color:#e67e22;font-weight:bold;'>未找到基本释义</div>`;
        if (machineTranslation) {
            html += `<div style='color:${color};margin:8px 0 0 0;font-weight:bold;'>***${machineTranslation}***</div>`;
        }
    }

    // 相关短语
    if (w && w.phrase && w.phrase.length > 0) {
        html += `<div style="margin-top:8px;"><b>相关短语：</b><ul>`;
        w.phrase.forEach((p: any) => {
            html += `<li>${p.pcontent}：${p.trans}</li>`;
        });
        html += `</ul></div>`;
    }

    // 例句
    let exampleList: any[] = [];
    if (w && w.exam_sents && w.exam_sents.length > 0) {
        exampleList = w.exam_sents.map((s: any) => ({
            eng: s.eng,
            chn: s.chn
        }));
    }
    // media_sents_part
    if (dictData.media_sents_part && Array.isArray(dictData.media_sents_part.sent) && dictData.media_sents_part.sent.length > 0) {
        exampleList = exampleList.concat(
            dictData.media_sents_part.sent.map((s: any) => ({
                eng: s.eng,
                chn: s.chn
            }))
        );
    }
    // blc
    if (dictData.blc && dictData.blc.blc_sents && dictData.blc.blc_sents.length > 0) {
        exampleList = exampleList.concat(
            dictData.blc.blc_sents.map((s: any) => ({
                eng: s.eng,
                chn: s.chn
            }))
        );
    }
    // collins 深度递归
    if (dictData.collins && dictData.collins.collins_entries && dictData.collins.collins_entries.length > 0) {
        dictData.collins.collins_entries.forEach((collinsEntry: any) => {
            if (collinsEntry.entries && collinsEntry.entries.entry) {
                collinsEntry.entries.entry.forEach((entry: any) => {
                    if (entry.tran_entry) {
                        entry.tran_entry.forEach((tran: any) => {
                            if (tran.exam_sents && tran.exam_sents.sent) {
                                tran.exam_sents.sent.forEach((sent: any) => {
                                    exampleList.push({
                                        eng: sent.eng_sent || sent.eng || '',
                                        chn: sent.chn_sent || sent.chn || ''
                                    });
                                });
                            }
                        });
                    }
                });
            }
        });
    }

    if (exampleList.length > 0) {
        html += `<div style="margin-top:12px;"><b>例句：</b><ul style="padding-left:1.5em;">`;
        exampleList.forEach((s: any) => {
            // 只显示中英双全的例句
            if (s.eng && s.chn) {
                html += `<li style="margin-bottom:1.2em;">
                    <div style="color:${color};font-weight:bold;font-size:1.05em;line-height:1.6;">${s.eng}</div>
                    <div style="margin-left:2em;color:#444;font-size:0.98em;line-height:1.6;margin-top:0.2em;">${s.chn}</div>
                </li>`;
            }
            // 其它情况（只有 eng 或只有 chn）一律不显示
        });
        html += `</ul></div>`;
    }

    // Collins词典
    if (dictData.collins && dictData.collins.entries && dictData.collins.entries.length > 0) {
        html += `<div style="margin-top:8px;"><b>Collins词典：</b><ul>`;
        dictData.collins.entries.forEach((entry: any) => {
            html += `<li>${entry.value || ''}</li>`;
        });
        html += `</ul></div>`;
    }

    // 英汉大词典
    if (dictData.bidec && dictData.bidec.word && dictData.bidec.word.length > 0) {
        const bw = dictData.bidec.word[0];
        html += `<div style="margin-top:8px;"><b>英汉大词典：</b><ul>`;
        (bw.trs||[]).forEach((tr: any)=> {
            html += `<li>${tr.tr[0].l.i[0]}</li>`;
        });
        html += `</ul></div>`;
    }

    // 网络释义
    if (dictData.web_trans && dictData.web_trans['web-translation'] && dictData.web_trans['web-translation'].length > 0) {
        html += `<div style="margin-top:8px;"><b>网络释义：</b><ul>`;
        dictData.web_trans['web-translation'].forEach((item: any) => {
            html += `<li>${item.key}: ${(item.trans || []).map((t: any) => t.value).join('; ')}</li>`;
        });
        html += `</ul></div>`;
    }

    return html;
}

function getFullAudioUrl(url: string): string {
    if (!url) return '';
    if (url.startsWith('http')) return url;
    if (url.startsWith('/')) return 'https://dict.youdao.com' + url;
    if (url.startsWith('dictvoice?') || url.startsWith('ttsvoice?')) return 'https://dict.youdao.com/' + url;
    if (url.startsWith('audio=')) return 'https://dict.youdao.com/dictvoice?' + url;
    // 兼容 apple&type=1 这种情况
    if (/^[^&]+&type=\d$/.test(url)) return `https://dict.youdao.com/dictvoice?audio=${url}`;
    return url;
}

// 抽取主逻辑
async function handleDictOrTranslate(q: string, app: App, settings: YoudaoSettings, plugin: any) {
    if (hasMinimizedModal()) {
        console.log('[handleDictOrTranslate] return: hasMinimizedModal, q:', q);
        new Notice("请先还原/关闭顶层弹窗，再进行其他操作");
        return;
    }
    console.log('[YoudaoPlugin] handleDictOrTranslate 入参:', q);
    function isChinese(text: string): boolean {
        return /[\u4e00-\u9fa5]/.test(text);
    }
    let targetLang = settings.textTargetLang;
    const lines = q.split(/\r?\n/).filter(line => line.trim().length > 0);
    // === 新增：复合翻译模式 ===
    if (settings.translateModel === 'composite') {
        // 多行或有句号，直接走微软批量翻译
        if (lines.length > 1 || (lines.length === 1 && /[。！？.!?]/.test(lines[0]))) {
            console.log('[handleDictOrTranslate] return: composite 多行/有句号，直接走微软批量翻译, q:', q);
            await unifiedBatchTranslateForMicrosoft({
                app,
                items: lines,
                getTargetLang: settings.isBiDirection
                    ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans'
                    : () => settings.textTargetLang === 'zh' ? 'zh-Hans' : settings.textTargetLang,
                msAdapter: new MicrosoftAdapter(
                    settings.microsoftKey || '',
                    settings.microsoftRegion || '',
                    settings.microsoftEndpoint || '',
                    app
                ),
                color: settings.textTranslationColor,
                mode: settings.textTranslationMode,
                sleepInterval: settings.microsoftSleepInterval ?? 250,
                plugin
            });
            return;
        }
        // 单行，判断分界
        const isZh = isChinese(q);
        let useDict = false;
        if (isZh) {
            const zhCount = (q.match(/[\u4e00-\u9fa5]/g) || []).length;
            const zhLimit = settings.dictWordLimitZh ?? 4;
            useDict = zhCount > 0 && zhCount <= zhLimit;
        } else {
            const wordCount = q.trim().split(/\s+/).length;
            const enLimit = settings.dictWordLimitEn ?? 3;
            useDict = wordCount > 0 && wordCount <= enLimit;
        }
        if (useDict) {
            try {
                const resp = await fetch(`http://127.0.0.1:${settings.serverPort}/api/dict?q=${encodeURIComponent(q)}`);
                if (!resp.ok) throw new Error('本地词典API请求失败');
                const dictData = await resp.json();
                // 判断是否有基本释义
                let hasBasic = false;
                if (Array.isArray(dictData.definitions) && dictData.definitions.length > 0) {
                    hasBasic = true;
                }
                let machineTranslation = '';
                if (!hasBasic) {
                    // 机器翻译目标语言逻辑（与超分界一致）
                    let tLang = 'en';
                    if (settings.isBiDirection) {
                        tLang = isChinese(q) ? 'en' : 'zh-Hans';
                    } else {
                        tLang = settings.textTargetLang === 'zh' ? 'zh-Hans' : settings.textTargetLang;
                    }
                    const from = isChinese(q) ? 'zh-Hans' : 'en';
                    machineTranslation = await new MicrosoftAdapter(
                        settings.microsoftKey || '',
                        settings.microsoftRegion || '',
                        settings.microsoftEndpoint || '',
                        app
                    ).translateText(q, from, tLang) || '';
                }
                const html = renderDictResult({
                    definitions: dictData.definitions || [],
                    examples: dictData.examples || [],
                    phrases: dictData.phrases || []
                }, isZh ? 'zh' : 'en', settings.textTranslationColor, settings.serverPort, machineTranslation, q);
                new DictResultModal(app, html, plugin).open();
                console.log('[handleDictOrTranslate] return: composite useDict, q:', q, 'mainLang:', isZh ? 'zh' : 'en', 'targetLang:', isZh ? 'en' : 'zh-Hans');
                return;
            } catch (e) {
                new Notice('本地词典释义获取失败：' + e.message);
                console.log('[handleDictOrTranslate] return: composite useDict error, q:', q, 'mainLang:', isZh ? 'zh' : 'en', 'targetLang:', isZh ? 'en' : 'zh-Hans', 'error:', e);
                return;
            }
        }
        // 超分界，走微软机器翻译
        let tLang = 'en';
        if (settings.isBiDirection) {
            tLang = isChinese(q) ? 'en' : 'zh-Hans';
        } else {
            tLang = settings.textTargetLang === 'zh' ? 'zh-Hans' : settings.textTargetLang;
        }
        if ((tLang === 'zh-Hans' && isChinese(q)) || (tLang === 'en' && !isChinese(q))) {
            new TranslateResultModal(app, q, q, settings.textTranslationColor, settings.textTranslationMode, plugin).open();
            console.log('[handleDictOrTranslate] return: composite 主语言和目标语言一致, q:', q, 'mainLang:', isZh ? 'zh' : 'en', 'targetLang:', tLang);
            return;
        }
        const from = isChinese(q) ? 'zh-Hans' : 'en';
        const translated = await new MicrosoftAdapter(
            settings.microsoftKey || '',
            settings.microsoftRegion || '',
            settings.microsoftEndpoint || '',
            app
        ).translateText(q, from, tLang);
        if (translated) {
            new TranslateResultModal(app, q, translated, settings.textTranslationColor, settings.textTranslationMode, plugin).open();
            console.log('[handleDictOrTranslate] return: composite 机器翻译成功, q:', q, 'mainLang:', isZh ? 'zh' : 'en', 'targetLang:', tLang);
        } else {
            new Notice('微软翻译失败');
            console.log('[handleDictOrTranslate] return: composite 机器翻译失败, q:', q, 'mainLang:', isZh ? 'zh' : 'en', 'targetLang:', tLang);
        }
        return;
    }
    // === 原有 youdao 模式 ===
    // --- 统一：无论 isBiDirection 开关状态，均分行分别翻译 ---
    if (lines.length > 1 || (lines.length === 1 && /[。！？.!?]/.test(lines[0]))) {
        console.log('[handleDictOrTranslate] return: youdao 多行/有句号，直接走批量翻译, q:', q);
        await unifiedBatchTranslate({
            app,
            items: lines,
            getTargetLang: settings.isBiDirection
                ? (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-CHS'
                : () => settings.textTargetLang,
            textAppKey: settings.textAppKey,
            textAppSecret: settings.textAppSecret,
            color: settings.textTranslationColor,
            mode: settings.textTranslationMode,
            port: settings.serverPort,
            sleepInterval: settings.sleepInterval ?? 250,
            plugin,
            isBiDirection: settings.isBiDirection
        });
        return;
    }
    // 单行文本，继续走原有逻辑
    // === 修改主语言判断方式为 detectMainLangSmart ===
    const mainLang = detectMainLangSmart(q);
    const isZh = mainLang === 'zh';
    if (settings.isBiDirection) {
        targetLang = isZh ? 'en' : 'zh-CHS';
    }
    if (targetLang === "zh-CHS" && isZh) {
        new TranslateResultModal(app, q, q, settings.textTranslationColor, settings.textTranslationMode, plugin).open();
        console.log('[handleDictOrTranslate] return: youdao 主语言和目标语言一致, q:', q, 'mainLang:', mainLang, 'targetLang:', targetLang);
        return;
    }
    let useDict = false;
    if (isZh) {
        const zhCount = (q.match(/[\u4e00-\u9fa5]/g) || []).length;
        const zhLimit = settings.dictWordLimitZh ?? 4;
        useDict = zhCount > 0 && zhCount <= zhLimit;
    } else {
        const wordCount = q.trim().split(/\s+/).length;
        const enLimit = settings.dictWordLimitEn ?? 3;
        useDict = wordCount > 0 && wordCount <= enLimit;
    }
    if (useDict) {
        try {
            const translator: TranslateService = new YoudaoAdapter(
                settings.textAppKey,
                settings.textAppSecret,
                app,
                settings.serverPort
            );
            const dictData = await translator.lookupWord?.(q.trim(), settings.serverPort);
            // 判断是否有基本释义
            let hasBasic = false;
            let w = null;
            if (dictData.ec && dictData.ec.word && dictData.ec.word.length > 0) w = dictData.ec.word[0];
            else if (dictData.ce && dictData.ce.word && dictData.ce.word.length > 0) w = dictData.ce.word[0];
            else if (dictData.ee && dictData.ee.word && dictData.ee.word.length > 0) w = dictData.ee.word[0];
            let trsList = [];
            if (w && w.trs && w.trs.length > 0) trsList = w.trs;
            else if (dictData.ec && dictData.ec.word && dictData.ec.word[0] && dictData.ec.word[0].trs && dictData.ec.word[0].trs.length > 0) trsList = dictData.ec.word[0].trs;
            else if (dictData.ce && dictData.ce.word && dictData.ce.word[0] && dictData.ce.word[0].trs && dictData.ce.word[0].trs.length > 0) trsList = dictData.ce.word[0].trs;
            else if (dictData.ee && dictData.ee.word && dictData.ee.word[0] && dictData.ee.word[0].trs && dictData.ee.word[0].trs.length > 0) trsList = dictData.ee.word[0].trs;
            if (trsList.length > 0) hasBasic = true;
            let machineTranslation = '';
            if (!hasBasic) {
                // 机器翻译目标语言逻辑
                let targetLang = settings.textTargetLang;
                if (settings.isBiDirection) {
                    const zhCount = (q.match(/[\u4e00-\u9fa5]/g) || []).length;
                    const enCount = (q.match(/[a-zA-Z]/g) || []).length;
                    targetLang = zhCount >= enCount ? 'en' : 'zh-CHS';
                }
                machineTranslation = await translator.translateText(q, 'auto', targetLang) || '';
            }
            const html = renderDictResult(dictData, isZh ? 'zh' : 'en', settings.textTranslationColor, settings.serverPort, machineTranslation, q);
            new DictResultModal(app, html, plugin).open();
            console.log('[handleDictOrTranslate] return: youdao useDict, q:', q, 'mainLang:', mainLang, 'targetLang:', targetLang);
            return;
        } catch (e) {
            new Notice('词典释义获取失败：' + e.message);
            console.log('[handleDictOrTranslate] return: youdao useDict error, q:', q, 'mainLang:', mainLang, 'targetLang:', targetLang, 'error:', e);
        }
    } else {
        // 新增：通过接口调用
        const translator: TranslateService = new YoudaoAdapter(
            settings.textAppKey,
            settings.textAppSecret,
            app,
            settings.serverPort
        );
        const translated = await translator.translateText(
            q,
            'auto',
            targetLang
        );
        if (translated) {
            new TranslateResultModal(app, q, translated, settings.textTranslationColor, settings.textTranslationMode, plugin).open();
            console.log('[handleDictOrTranslate] return: youdao 机器翻译成功, q:', q, 'mainLang:', mainLang, 'targetLang:', targetLang);
        } else {
            new Notice("翻译失败");
            console.log('[handleDictOrTranslate] return: youdao 机器翻译失败, q:', q, 'mainLang:', mainLang, 'targetLang:', targetLang);
        }
        return;
    }
}

function getFirstHotkeyCombo(app: App, commandId: string) {
    const hotkeyManager = (app as any).hotkeyManager;
    if (!hotkeyManager || !hotkeyManager.customKeys) {
        console.log('[YoudaoPlugin] hotkeyManager/customKeys 不可用');
        return null;
    }
    const idPattern = new RegExp(`(^|[ :])${commandId}$`);
    let foundKey = null;
    for (const k in hotkeyManager.customKeys) {
        if (idPattern.test(k)) {
            foundKey = k;
            console.log('[YoudaoPlugin] 匹配到 customKeys key:', k, hotkeyManager.customKeys[k]);
            break;
        }
    }
    let custom = foundKey ? hotkeyManager.customKeys[foundKey] : [];
    if ((!custom || custom.length === 0) && Array.isArray(hotkeyManager.bakedHotkeys)) {
        for (const entry of hotkeyManager.bakedHotkeys) {
            if (entry && entry.id && idPattern.test(entry.id) && Array.isArray(entry.keys) && entry.keys.length > 0) {
                custom = [entry.keys[0]];
                console.log('[YoudaoPlugin] 匹配到 bakedHotkeys:', entry);
                break;
            }
        }
    }
    if (!custom || custom.length === 0) {
        console.log('[YoudaoPlugin] 没有匹配到任何快捷键');
        return null;
    }
    const h = custom[0];
    // 兼容 ctrlKey/shiftKey/altKey 字段
    const ctrl = (h.modifiers || []).includes('Ctrl') || h.ctrlKey === true;
    const shift = (h.modifiers || []).includes('Shift') || h.shiftKey === true;
    const alt = (h.modifiers || []).includes('Alt') || h.altKey === true;
    console.log('[YoudaoPlugin] 返回快捷键信息:', { ctrl, shift, alt, key: h.key });
    console.log('[YoudaoPlugin] customKeys value:', custom);
    return {
        
        ctrl,
        shift,
        alt,
        key: h.key?.toLowerCase()
    };
}

function splitToSentences(text: string): string[] {
    // 按中英文句号、问号、感叹号、换行等切分
    return text.split(/(?<=[。！？.!?\\n])/).map(s => s.trim()).filter(Boolean);
}

function chunkSentences(sentences: string[], maxLen: number): string[] {
    const chunks: string[] = [];
    let buffer = '';
    for (const sent of sentences) {
        if ((buffer + sent).length > maxLen) {
            if (buffer) chunks.push(buffer);
            buffer = sent;
        } else {
            buffer += sent;
        }
    }
    if (buffer) chunks.push(buffer);
    return chunks;
}

function sleep(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
}

async function robustTranslateChunk(chunk: string, textAppKey: string, textAppSecret: string, targetLang: string, app: App, port: string, maxRetry = 5): Promise<string> {
    let lastErr = null;
    // 新增：接口化
    const translator: TranslateService = new YoudaoAdapter(textAppKey, textAppSecret, app, port);
    for (let i = 0; i < maxRetry; i++) {
        const t = await translator.translateText(chunk, 'auto', targetLang);
        console.log('[robustTranslateChunk] 调用机器翻译API, chunk:', chunk, 'targetLang:', targetLang, '返回:', t, '第', i+1, '次尝试');
        if (t && t.trim()) return t; // 只要有返回就直接用
        lastErr = t;
        await sleep(250); // 或120
    }
    console.log('[robustTranslateChunk] 翻译失败, chunk:', chunk, 'targetLang:', targetLang, 'lastErr:', lastErr);
    return ''; // 失败返回空串
}

// 全局最小化弹窗检查工具
function hasMinimizedModal() {
    return globalMinimizedModals && globalMinimizedModals.length > 0;
}

// 新增：统一的分行/分块批量翻译逻辑
async function unifiedBatchTranslate({
    app,
    items,
    getTargetLang,
    textAppKey,
    textAppSecret,
    color,
    mode,
    port,
    sleepInterval,
    plugin,
    isBiDirection
}: {
    app: App,
    items: string[],
    getTargetLang: (text: string) => string,
    textAppKey: string,
    textAppSecret: string,
    color: string,
    mode: string,
    port: string,
    sleepInterval: number,
    plugin: any,
    isBiDirection: boolean
}) {
    let loading = document.createElement('div');
    loading.className = 'youdao-loading';
    loading.innerHTML = `<div class='youdao-spinner'></div> 翻译正在进行，请稍等...<div class='youdao-progress' style='margin-top:16px;font-size:1.1em;'></div>`;
    Object.assign(loading.style, {
      position: 'fixed', left: '50%', top: '30%', transform: 'translate(-50%, -50%)',
      zIndex: 99999, background: '#fff', padding: '32px 48px', borderRadius: '12px',
      boxShadow: '0 2px 16px #0002', fontSize: '1.2em', textAlign: 'center'
    });
    document.body.appendChild(loading);
    const progressEl = loading.querySelector('.youdao-progress');
    try {
        let translatedArr: string[] = [];
        let successCount = 0;
        for (let i = 0; i < items.length; i++) {
            const text = (items[i] || '').trim();
            if (!text) {
                translatedArr.push('');
                if (progressEl) progressEl.textContent = `已完成${i+1}/${items.length}`;
                continue;
            }
            // 优化跳过逻辑
            const mainLang = detectMainLangSmart(text);
            let tLang = getTargetLang(text);
            console.log('[unifiedBatchTranslate] 待翻译文本:', text, '目标语言:', tLang);
            // 关掉中英互译时，目标语言为中文且原文为中文，直接返回原文
            const isAllZh = /^[\u4e00-\u9fa5\s，。！？、；："'（）【】《》…—·]*$/.test(text);
            const isAllEn = /^[a-zA-Z0-9\s.,!?;:'"()\[\]{}<>@#$%^&*_+=|\\/-]*$/.test(text);
            if ((tLang.startsWith('zh') && mainLang === 'zh' && isAllZh) ||
                (tLang.startsWith('en') && mainLang === 'en' && isAllEn)) {
                translatedArr.push(text);
                successCount++;
                if (progressEl) progressEl.textContent = `已完成${i+1}/${items.length}`;
                await sleep(sleepInterval ?? 250);
                console.log('[unifiedBatchTranslate] 跳过翻译，原文与目标语言一致:', text);
                continue;
            }
            let translated = await robustTranslateChunk(text, textAppKey, textAppSecret, tLang, app, port, 5);
            // 关键逻辑：
            if (isBiDirection) {
                if (translated && translated.trim() && translated.trim() !== text) {
                    translatedArr.push(translated);
                    successCount++;
                    console.log('[unifiedBatchTranslate] 翻译成功:', text, '->', translated);
                } else {
                    translatedArr.push('（没有基本释义）');
                    console.log('[unifiedBatchTranslate] 翻译无基本释义:', text, '返回:', translated);
                }
            } else {
                // 关闭互译时，原文=译文也正常显示
                if (translated && translated.trim()) {
                    translatedArr.push(translated);
                    successCount++;
                    console.log('[unifiedBatchTranslate] 翻译成功:', text, '->', translated);
                } else {
                    translatedArr.push('（没有基本释义）');
                    console.log('[unifiedBatchTranslate] 翻译无基本释义:', text, '返回:', translated);
                }
            }
            if (progressEl) progressEl.textContent = `已完成${i+1}/${items.length}`;
            await sleep(sleepInterval ?? 250);
        }
        let original = items.join('\n');
        let translated = translatedArr.join('\n');
        new TranslateResultModal(app, original, translated, color, mode, plugin, successCount, items.length).open();
    } catch (e) {
        console.error('[unifiedBatchTranslate] 批量翻译异常:', e);
    } finally {
        if (loading) loading.remove();
    }
}

class VocabBookModal extends Modal {
    plugin: any;
    _isMinimized: boolean = false;
    _isMaximized: boolean = false;
    _modalId: string;
    _headerEl: HTMLElement | null = null;
    constructor(app: App, plugin: any) {
        super(app);
        this.plugin = plugin;
        this._modalId = getUniqueModalId();
    }
    onOpen() {
        globalVocabBookModal = this;
        const { contentEl, modalEl } = this;
        contentEl.empty();
        // 默认样式
        contentEl.style.padding = "24px";
        contentEl.style.background = "#f8fafc";
        contentEl.style.minWidth = "600px";
        contentEl.style.minHeight = "480px";
        contentEl.style.maxHeight = "80vh";
        contentEl.style.overflowY = "auto";

        // Cursor 风格搜索区
        const searchBoxWrapper = contentEl.createEl("div");
        searchBoxWrapper.style.position = "absolute";
        searchBoxWrapper.style.top = "24px";
        searchBoxWrapper.style.right = "32px";
        searchBoxWrapper.style.zIndex = "1000";
        searchBoxWrapper.style.display = "flex";
        searchBoxWrapper.style.alignItems = "center";
        searchBoxWrapper.style.background = "#fff";
        searchBoxWrapper.style.border = "1px solid #e0e0e0";
        searchBoxWrapper.style.borderRadius = "8px";
        searchBoxWrapper.style.boxShadow = "0 2px 8px #0001";
        searchBoxWrapper.style.padding = "4px 8px";
        searchBoxWrapper.style.gap = "0";

        // 新增：搜索模式切换按钮
        let searchMode: 'fuzzy' | 'word' = 'fuzzy';
        let performSearch: () => void;
        let modeBtn: HTMLButtonElement;
        modeBtn = searchBoxWrapper.createEl("button", { text: "模糊搜索" });
        modeBtn.style.marginRight = "8px";
        modeBtn.style.padding = "4px 10px";
        modeBtn.style.fontSize = "0.98em";
        modeBtn.style.border = "none";
        modeBtn.style.background = "#e3eafc";
        modeBtn.style.borderRadius = "4px";
        modeBtn.style.cursor = "pointer";
        // 先定义 modeBtn.onclick，后赋值 performSearch
        modeBtn.onclick = () => {
            if (searchMode === 'fuzzy') {
                searchMode = 'word';
                modeBtn.textContent = '单词名和短语名搜索';
            } else {
                searchMode = 'fuzzy';
                modeBtn.textContent = '模糊搜索';
            }
            if (performSearch) performSearch();
        };

        const searchInput = searchBoxWrapper.createEl("input");
        searchInput.type = "text";
        searchInput.placeholder = "输入单词搜索";
        searchInput.style.padding = "4px 8px";
        searchInput.style.fontSize = "1em";
        searchInput.style.border = "none";
        searchInput.style.outline = "none";
        searchInput.style.background = "transparent";

        const upBtn = searchBoxWrapper.createEl("button", { text: "↑" });
        upBtn.title = "上一个";
        upBtn.style.padding = "2px 8px";
        upBtn.style.margin = "0 2px";
        upBtn.style.border = "none";
        upBtn.style.background = "#f5f5f5";
        upBtn.style.borderRadius = "4px";
        upBtn.style.cursor = "pointer";

        const downBtn = searchBoxWrapper.createEl("button", { text: "↓" });
        downBtn.title = "下一个";
        downBtn.style.padding = "2px 8px";
        downBtn.style.margin = "0 2px";
        downBtn.style.border = "none";
        downBtn.style.background = "#f5f5f5";
        downBtn.style.borderRadius = "4px";
        downBtn.style.cursor = "pointer";

        const resultInfo = searchBoxWrapper.createEl("span");
        resultInfo.style.margin = "0 8px";
        resultInfo.style.fontSize = "0.98em";
        resultInfo.style.color = "#888";

        const searchBtn = searchBoxWrapper.createEl("button", { text: "搜索" });
        searchBtn.style.padding = "4px 12px";
        searchBtn.style.fontSize = "1em";
        searchBtn.style.marginLeft = "8px";
        searchBtn.style.border = "none";
        searchBtn.style.background = "#e3eafc";
        searchBtn.style.borderRadius = "4px";
        searchBtn.style.cursor = "pointer";

        contentEl.appendChild(searchBoxWrapper);

        // 分页变量
        let currentPage = 1;
        let pageSize = this._isMaximized ? 18 : 12; // 固定每页渲染数，最大化18，普通12

        // 标题栏和按钮
        const header = contentEl.createEl("div");
        header.style.display = "flex";
        header.style.justifyContent = "space-between";
        header.style.alignItems = "center";
        header.style.marginBottom = "16px";
        this._headerEl = header;
        const title = header.createEl("h2", { text: "我的词汇本" });
        title.style.marginBottom = "0";
        // 按钮组
        const btnGroup = header.createEl("div");
        btnGroup.style.display = "flex";
        btnGroup.style.gap = "12px";
        // 最小化按钮
        const minBtn = btnGroup.createEl("button", { text: "–" });
        minBtn.title = "最小化";
        minBtn.style.fontSize = "1.3em";
        minBtn.style.width = "32px";
        minBtn.style.height = "32px";
        minBtn.style.border = "none";
        minBtn.style.background = "none";
        minBtn.style.cursor = "pointer";
        minBtn.onclick = () => {
            this._isMinimized = true;
            modalEl.style.display = "none";
            // 任务栏区域
            let bar = document.getElementById('youdao-modal-taskbar');
            if (!bar) {
                bar = document.createElement('div');
                bar.id = 'youdao-modal-taskbar';
                bar.style.position = 'fixed';
                bar.style.left = '50%';
                bar.style.transform = 'translateX(-50%)';
                bar.style.right = '';
                bar.style.bottom = '0';
                bar.style.height = '44px';
                bar.style.background = 'rgba(255,255,255,0.95)';
                bar.style.zIndex = '99999';
                bar.style.display = 'flex';
                bar.style.alignItems = 'center';
                bar.style.gap = '12px';
                bar.style.padding = '0 16px';
                document.body.appendChild(bar);
            }
            // 还原按钮
            const restoreBtn = document.createElement('button');
            restoreBtn.textContent = '词汇本';
            restoreBtn.title = '词汇本';
            restoreBtn.style.margin = '0 8px';
            restoreBtn.style.padding = '6px 18px';
            restoreBtn.style.fontSize = '1em';
            restoreBtn.style.borderRadius = '6px';
            restoreBtn.style.background = '#f5f5f5';
            restoreBtn.style.border = '1px solid #ccc';
            restoreBtn.style.cursor = 'pointer';
            restoreBtn.onclick = () => {
                modalEl.style.display = '';
                this._isMinimized = false;
                restoreBtn.remove();
                globalMinimizedModals = globalMinimizedModals.filter(m => m.id !== this._modalId);
                if (bar && bar.children.length === 0) bar.remove();
            };
            bar.appendChild(restoreBtn);
            globalMinimizedModals.push({ id: this._modalId, restoreBtn, modal: this });
        };
        // 最大化按钮
        const maxBtn = btnGroup.createEl("button", { text: "☐" });
        maxBtn.title = "最大化";
        maxBtn.style.fontSize = "1.1em";
        maxBtn.style.width = "32px";
        maxBtn.style.height = "32px";
        maxBtn.style.border = "none";
        maxBtn.style.background = "none";
        maxBtn.style.cursor = "pointer";
        // 最大化样式应用函数
        let tableDiv: HTMLDivElement;
        let lastContainer: HTMLDivElement | null = null;
        const applyMaximizedStyle = () => {
            const ROW_HEIGHT = 38; // 每行高度
            const HEADER_HEIGHT = 48; // 表头高度
            pageSize = this._isMaximized ? 18 : 12;
            if (lastContainer) {
                lastContainer.style.height = (ROW_HEIGHT * pageSize + HEADER_HEIGHT) + 'px';
                lastContainer.style.maxHeight = (ROW_HEIGHT * pageSize + HEADER_HEIGHT) + 'px';
                lastContainer.style.overflowY = 'auto';
                lastContainer.style.overflowX = 'hidden';
            }


            modalEl.style.pointerEvents = 'auto';

            // if (this._isMaximized) {
            //     modalEl.style.width = '100vw';
            //     modalEl.style.height = '100vh';
            //     modalEl.style.left = '0';
            //     modalEl.style.top = '0';
            //     modalEl.style.maxWidth = '100vw';
            //     modalEl.style.maxHeight = '100vh';
            //     modalEl.style.borderRadius = '0';
            //     modalEl.style.zIndex = '999999';
            //     contentEl.style.width = '100vw';
            //     contentEl.style.height = '100vh';
            //     contentEl.style.minWidth = '0';
            //     contentEl.style.minHeight = '0';
            //     contentEl.style.maxWidth = '100vw';
            //     contentEl.style.maxHeight = '100vh';
            //     contentEl.style.overflow = 'hidden';
            //     if (tableDiv) {
            //         tableDiv.style.width = '100vw';
            //         tableDiv.style.height = 'calc(100vh - 120px)';
            //         tableDiv.style.maxWidth = '100vw';
            //         tableDiv.style.maxHeight = 'calc(100vh - 120px)';
            //         tableDiv.style.overflow = 'hidden';
            //     }
            //     if (lastContainer) {
            //         lastContainer.style.height = '100%';
            //         lastContainer.style.maxHeight = '100%';
            //     }
            // }
            // else {
            //     modalEl.style.width = '';
            //     modalEl.style.height = '';
            //     modalEl.style.left = '';
            //     modalEl.style.top = '';
            //     modalEl.style.maxWidth = '';
            //     modalEl.style.maxHeight = '';
            //     modalEl.style.borderRadius = '';
            //     modalEl.style.zIndex = String(Date.now());
            //     contentEl.style.width = '';
            //     contentEl.style.height = '';
            //     contentEl.style.minWidth = '600px';
            //     contentEl.style.minHeight = '480px';
            //     contentEl.style.maxWidth = '';
            //     contentEl.style.maxHeight = '80vh';
            //     contentEl.style.overflowY = 'auto';
            //     if (tableDiv) {
            //         tableDiv.style.width = '';
            //         tableDiv.style.height = '';
            //         tableDiv.style.maxWidth = '';
            //         tableDiv.style.maxHeight = '';
            //         tableDiv.style.overflow = '';
            //     }
            //     if (lastContainer) {
            //         lastContainer.style.maxHeight = '';
            //     }
            // }
        };


        maxBtn.onclick = () => {
            this._isMaximized = !this._isMaximized;
            if (this._isMaximized) {
                modalEl.style.position = 'fixed'; // 关键：固定定位
                modalEl.style.left = '0';
                modalEl.style.top = '0';
                modalEl.style.width = '100vw';
                modalEl.style.height = '100vh';
                modalEl.style.maxWidth = '100vw';
                modalEl.style.maxHeight = '100vh';
                modalEl.style.borderRadius = '0';
                modalEl.style.zIndex = '999999';
                contentEl.style.width = '100vw';
                contentEl.style.height = '100vh';
                contentEl.style.minWidth = '0';
                contentEl.style.minHeight = '0';
                contentEl.style.maxWidth = '100vw';
                contentEl.style.maxHeight = '100vh';
                contentEl.style.overflow = 'hidden';
            } else {
                modalEl.style.position = '';
                modalEl.style.left = '';
                modalEl.style.top = '';
                modalEl.style.width = '';
                modalEl.style.height = '';
                modalEl.style.maxWidth = '';
                modalEl.style.maxHeight = '';
                modalEl.style.borderRadius = '';
                modalEl.style.zIndex = String(Date.now());
                contentEl.style.width = '';
                contentEl.style.height = '';
                contentEl.style.minWidth = '600px';
                contentEl.style.minHeight = '480px';
                contentEl.style.maxWidth = '';
                contentEl.style.maxHeight = '80vh';
                contentEl.style.overflowY = 'auto';
            }
            applyMaximizedStyle();
            maxBtn.style.fontWeight = this._isMaximized ? 'bold' : '';
            renderTable();
        };

        // 功能按钮区
        const btnBar = contentEl.createEl("div");
        btnBar.style.display = "flex";
        btnBar.style.gap = "12px";
        btnBar.style.marginBottom = "18px";

        // 导出按钮
        const exportBtn = btnBar.createEl("button", { text: "导出" });
        exportBtn.onclick = () => exportVocabBook(this.plugin);

        // 导入按钮
        const importBtn = btnBar.createEl("button", { text: "导入" });
        importBtn.onclick = () => {
            const input = document.createElement("input");
            input.type = "file";
            input.accept = ".json";
            input.onchange = async (e: any) => {
                if (input.files && input.files.length) {
                    await importVocabBook(this.plugin, input.files[0]);
                    this.onOpen(); // 刷新
                }
            };
            input.click();
        };

        // 回收站按钮
        const trashBtn = btnBar.createEl("button", { text: "回收站" });
        trashBtn.onclick = () => {
            const trashModal = new Modal(this.app);
            trashModal.contentEl.style.padding = "24px";
            trashModal.contentEl.style.background = "#f8fafc";
            trashModal.contentEl.style.minWidth = "600px";
            trashModal.contentEl.style.maxHeight = "80vh";
            trashModal.contentEl.style.overflowY = "auto";
            trashModal.contentEl.createEl("h2", { text: "回收站" });
            // === 清空回收站按钮 ===
            const clearBtn = trashModal.contentEl.createEl("button", { text: "清空回收站" });
            clearBtn.style.background = "#e53935";
            clearBtn.style.color = "#fff";
            clearBtn.style.padding = "6px 18px";
            clearBtn.style.borderRadius = "6px";
            clearBtn.style.fontWeight = "bold";
            clearBtn.style.marginBottom = "18px";
            clearBtn.style.marginRight = "12px";
            clearBtn.onclick = () => {
                if (!confirm("确定要清空回收站吗？此操作不可恢复。")) return;
                if (!confirm("再次确认：清空回收站将永久删除所有内容，确定继续？")) return;
                this.plugin.trashData.vocabBookTrash = [];
                setTimeout(() => { this.plugin.saveTrashData(); }, 0);
                new Notice("回收站已清空");
                trashModal.close();
            };
            const trash = this.plugin.trashData.vocabBookTrash || [];
            if (!trash.length) {
                trashModal.contentEl.createEl("div", { text: "回收站为空" }).style.color = "#888";
            } else {
                // 创建表格布局容器
                const tableContainer = trashModal.contentEl.createEl("div");
                tableContainer.style.display = "table";
                tableContainer.style.width = "100%";
                tableContainer.style.borderCollapse = "collapse";
                
                // 表头
                const headerRow = tableContainer.createEl("div");
                headerRow.style.display = "table-row";
                headerRow.style.fontWeight = "bold";
                headerRow.style.backgroundColor = "#f5f5f5";
                headerRow.style.borderBottom = "2px solid #ddd";
                
                const headerWord = headerRow.createEl("div");
                headerWord.style.display = "table-cell";
                headerWord.style.padding = "8px";
                headerWord.style.borderRight = "1px solid #ddd";
                headerWord.textContent = "单词";
                
                const headerTranslation = headerRow.createEl("div");
                headerTranslation.style.display = "table-cell";
                headerTranslation.style.padding = "8px";
                headerTranslation.style.borderRight = "1px solid #ddd";
                headerTranslation.textContent = "释义";
                
                const headerActions = headerRow.createEl("div");
                headerActions.style.display = "table-cell";
                headerActions.style.padding = "8px";
                headerActions.textContent = "操作";
                
                trash.forEach((v: any, idx: number) => {
                    const row = tableContainer.createEl("div");
                    row.style.display = "table-row";
                    row.style.borderBottom = "1px solid #eee";
                    
                    const wordCell = row.createEl("div");
                    wordCell.style.display = "table-cell";
                    wordCell.style.padding = "8px";
                    wordCell.style.borderRight = "1px solid #ddd";
                    wordCell.style.fontWeight = "bold";
                    wordCell.textContent = v.word;
                    
                    const translationCell = row.createEl("div");
                    translationCell.style.display = "table-cell";
                    translationCell.style.padding = "8px";
                    translationCell.style.borderRight = "1px solid #ddd";
                    translationCell.textContent = v.translation;
                    
                    const actionsCell = row.createEl("div");
                    actionsCell.style.display = "table-cell";
                    actionsCell.style.padding = "8px";
                    actionsCell.style.textAlign = "center";
                    
                    // 按钮横向排列
                    const btnContainer = actionsCell.createEl("div");
                    btnContainer.style.display = "flex";
                    btnContainer.style.gap = "8px";
                    btnContainer.style.justifyContent = "center";
                    
                    // 恢复按钮
                    const restoreBtn = btnContainer.createEl("button", { text: "恢复" });
                    restoreBtn.style.padding = "4px 12px";
                    restoreBtn.style.borderRadius = "4px";
                    restoreBtn.style.border = "1px solid #4caf50";
                    restoreBtn.style.background = "#4caf50";
                    restoreBtn.style.color = "white";
                    restoreBtn.style.cursor = "pointer";
                    restoreBtn.onclick = () => {
                        this.plugin.vocabData.vocabBook.push(v);
                        this.plugin.trashData.vocabBookTrash.splice(idx, 1);
                        setTimeout(() => { this.plugin.saveVocabData(); this.plugin.saveTrashData(); }, 0);
                        trashModal.close();
                        renderTable();
                    };
                    
                    // 彻底删除按钮
                    const delBtn = btnContainer.createEl("button", { text: "彻底删除" });
                    delBtn.style.padding = "4px 12px";
                    delBtn.style.borderRadius = "4px";
                    delBtn.style.border = "1px solid #f44336";
                    delBtn.style.background = "#f44336";
                    delBtn.style.color = "white";
                    delBtn.style.cursor = "pointer";
                    delBtn.onclick = () => {
                        this.plugin.trashData.vocabBookTrash.splice(idx, 1);
                        setTimeout(() => { this.plugin.saveTrashData(); }, 0);
                        trashModal.close();
                        renderTable();
                    };
                });
            }
            trashModal.open();
        };

        // 新建单词/短语按钮
        const createBtn = btnBar.createEl("button", { text: "新建单词/短语" });
        createBtn.onclick = () => {
            // 插入到词汇本最前面
            const newItem = {
                word: '',
                translation: '',
                example: '',
                group: '',
                notes: '',
                mastered: false,
                addedAt: Date.now()
            };
            this.plugin.vocabData.vocabBook.unshift(newItem);
            setTimeout(() => { this.plugin.saveVocabData(); }, 0);
            // 重新渲染表格，跳转到第一页
            currentPage = 1;
            renderTable();
            // 自动滚动到新行
            setTimeout(() => {
                const trs = tableDiv.querySelectorAll('tr');
                if (trs[0]) trs[0].scrollIntoView({ behavior: 'smooth', block: 'center' });
            }, 100);
        };

        // 词汇本表格
        tableDiv = contentEl.createEl("div");
        tableDiv.style.marginTop = "12px";

        // 渲染表格
        const renderTable = () => {
            tableDiv.empty();
            // 只用 filteredData 参与分页和渲染
            let data = filteredData && Array.isArray(filteredData) ? filteredData : [];
            // 分页
            const totalPages = Math.max(1, Math.ceil(data.length / pageSize));
            if (currentPage > totalPages) currentPage = totalPages;
            if (currentPage < 1) currentPage = 1;
            const start = (currentPage - 1) * pageSize;
            const end = Math.min(start + pageSize, data.length);
            const pageData = data.slice(start, end);
            console.log('[YoudaoPlugin] renderTable data.length:', data.length, 'totalPages:', totalPages, 'currentPage:', currentPage);
            // 表头
            const table = tableDiv.createEl("table");
            table.style.width = "100%";
            table.style.background = "#fff";
            table.style.borderRadius = "0";
            table.style.overflow = "visible";
            table.style.boxShadow = "none";
            table.style.margin = "0";
            table.style.tableLayout = "fixed";
            table.style.height = "100%";
            const thead = table.createEl("thead");
            thead.innerHTML = "<tr><th>单词</th><th>释义</th><th>例句</th><th>分组</th><th>备注</th><th>操作</th></tr>";

            // container高度自适应，最大化/非最大化分别设置maxHeight
            const container = tableDiv.createEl("div");
            lastContainer = container;
            container.style.height = '';
            if (this._isMaximized) {
                container.style.maxHeight = 'calc(100vh - 120px)';
            } else {
                container.style.maxHeight = '80vh';
            }
            container.style.width = '100%';
            container.style.maxWidth = '100%';
            container.style.overflowY = "auto";
            container.style.overflowX = "hidden";
            container.style.position = "relative";
            container.style.background = "#fff";
            container.style.borderRadius = "0";
            container.style.boxShadow = "none";
            container.style.marginTop = "0";
            container.style.display = "flex";
            container.style.flexDirection = "column";
            container.style.alignItems = "stretch";

            // 创建表格 tbody
            let tbody: HTMLTableSectionElement | null = null;
            const self = this;
            function renderRows() {
                if (tbody) tbody.remove();
                tbody = document.createElement("tbody");
                // 只渲染当前页 pageData
                for (let i = 0; i < pageData.length; i++) {
                    const v = pageData[i];
                    const tr = document.createElement("tr");
                    // 行高亮功能：只允许一行高亮
                    tr.onclick = function(e) {
                        if (e.target instanceof HTMLElement && e.target.tagName === 'BUTTON') return;
                        // 先移除所有高亮
                        const allRows = tableDiv.querySelectorAll('tr');
                        allRows.forEach(row => {
                            row.classList.remove('vocab-row-highlight');
                            (row as HTMLElement).style.background = '';
                        });
                        // 如果本行已高亮则取消，否则高亮
                        if (tr.classList.contains('vocab-row-highlight')) {
                            tr.classList.remove('vocab-row-highlight');
                            tr.style.background = '';
                        } else {
                            tr.classList.add('vocab-row-highlight');
                            tr.style.background = '#ffe082';
                        }
                    };
                    // 新增：右键菜单插入新单词/短语
                    tr.oncontextmenu = (e) => {
                        e.preventDefault();
                        // 移除已有菜单
                        const oldMenu = document.getElementById('vocab-context-menu');
                        if (oldMenu) oldMenu.remove();
                        const menu = document.createElement('div');
                        menu.id = 'vocab-context-menu';
                        menu.style.position = 'fixed';
                        menu.style.left = e.clientX + 'px';
                        menu.style.top = e.clientY + 'px';
                        menu.style.background = '#fff';
                        menu.style.border = '1px solid #e0e0e0';
                        menu.style.borderRadius = '6px';
                        menu.style.boxShadow = '0 2px 8px #0002';
                        menu.style.zIndex = '999999';
                        menu.style.padding = '6px 0';
                        menu.style.minWidth = '220px';
                        menu.style.fontSize = '1em';
                        // 插入新单词/短语选项
                        const insertItem = document.createElement('div');
                        insertItem.textContent = '在该单词/短语下方创建一个新单词/短语';
                        insertItem.style.padding = '8px 18px';
                        insertItem.style.cursor = 'pointer';
                        insertItem.onmouseenter = () => insertItem.style.background = '#e3eafc';
                        insertItem.onmouseleave = () => insertItem.style.background = '';
                        insertItem.onclick = () => {
                            menu.remove();
                            // 计算全局插入位置
                            const globalIdx = (currentPage - 1) * pageSize + i;
                            // 插入新空白对象
                            const newItem = {
                                word: '',
                                translation: '',
                                example: '',
                                group: '',
                                notes: '',
                                mastered: false,
                                addedAt: Date.now()
                            };
                            // 插入到全局 filteredData 的正确位置
                            const allData = self.plugin.vocabData.vocabBook;
                            // 找到 filteredData[globalIdx] 在 allData 的真实索引
                            const refItem = filteredData[globalIdx];
                            const realIdx = allData.indexOf(refItem);
                            if (realIdx !== -1) {
                                allData.splice(realIdx + 1, 0, newItem);
                                setTimeout(() => { self.plugin.saveVocabData(); }, 0);
                                renderTable();
                                // 自动滚动到新行
                                setTimeout(() => {
                                    const trs = tableDiv.querySelectorAll('tr');
                                    if (trs[globalIdx + 1]) {
                                        trs[globalIdx + 1].scrollIntoView({ behavior: 'smooth', block: 'center' });
                                    }
                                }, 100);
                            }
                        };
                        menu.appendChild(insertItem);
                        // 点击其他地方关闭菜单
                        setTimeout(() => {
                            const closeMenu = (ev: MouseEvent) => {
                                if (!menu.contains(ev.target as Node)) menu.remove();
                            };
                            document.addEventListener('mousedown', closeMenu, { once: true });
                        }, 10);
                        document.body.appendChild(menu);
                    };
                    // 搜索高亮（只高亮当前 searchMatches[searchIndex] 指向的那一行，优先级高于手动高亮）
                    if (searchMatches.length > 0 && highlightActive) {
                        const globalIndex = (currentPage - 1) * pageSize + i;
                        if (searchMatches[searchIndex] === globalIndex) {
                            tr.style.background = '#ffe082';
                            tr.style.border = '2px solid #1a73e8';
                        } else {
                            tr.style.border = '';
                        }
                    }
                    const tdWord = document.createElement("td");
                    if (!v.word) {
                        const input = document.createElement('input');
                        input.type = 'text';
                        input.placeholder = '单词/短语';
                        input.value = v.word;
                        input.style.width = '98%';
                        input.onchange = () => { v.word = input.value; setTimeout(() => { self.plugin.saveVocabData(); }, 0); };
                        tdWord.appendChild(input);
                    } else {
                    tdWord.textContent = v.word;
                    }
                    tr.appendChild(tdWord);
                    const tdTrans = document.createElement("td");
                    tdTrans.style.fontWeight = 'bold';
                    tdTrans.style.whiteSpace = 'pre-line';
                    tdTrans.style.fontSize = '1em';
                    // 释义列全部用 textarea 可编辑，支持多行
                    const transTextarea = document.createElement('textarea');
                    transTextarea.style.width = '98%';
                    transTextarea.style.minHeight = '48px';
                    transTextarea.style.fontSize = '1em';
                    transTextarea.style.fontWeight = 'bold';
                    transTextarea.style.background = 'transparent';
                    transTextarea.style.border = 'none';
                    transTextarea.style.resize = 'vertical';
                    transTextarea.value = v.translation || '';
                    transTextarea.onchange = () => {
                        v.translation = transTextarea.value;
                        setTimeout(() => { self.plugin.saveVocabData(); }, 0);
                    };
                    tdTrans.appendChild(transTextarea);
                    tr.appendChild(tdTrans);
                    const tdEx = document.createElement("td");
                    const exBtn = document.createElement('button');
                    exBtn.textContent = '例句';
                    exBtn.onclick = () => {
                        const modal = new Modal(self.app);
                        modal.contentEl.style.padding = "24px";
                        modal.contentEl.createEl("h2", { text: v.word + ' 例句' });
                        const textarea = document.createElement('textarea');
                        textarea.style.width = '100%';
                        textarea.style.height = '120px';
                        textarea.value = v.example || '';
                        modal.contentEl.appendChild(textarea);
                        const saveBtn = document.createElement('button');
                        saveBtn.textContent = '保存';
                        saveBtn.style.marginTop = '16px';
                        saveBtn.onclick = () => {
                            v.example = textarea.value;
                            setTimeout(() => { self.plugin.saveVocabData(); }, 0);
                            modal.close();
                            renderTable();
                        };
                        modal.contentEl.appendChild(saveBtn);
                        modal.open();
                    };
                    tdEx.appendChild(exBtn);
                    tr.appendChild(tdEx);
                    const tdGroup = document.createElement("td");
                    let groupArr: any[] = [];
                    if (Array.isArray(v.group)) groupArr = v.group;
                    else if (typeof v.group === 'string' && v.group) groupArr = [v.group];
                    if (groupArr.length > 0) {
                        tdGroup.innerHTML = groupArr.map(path => {
                            if (Array.isArray(path)) return `<div>${path.join('/')}</div>`;
                            if (typeof path === 'string') return `<div>${path}</div>`;
                            return '';
                        }).join('');
                    } else {
                        tdGroup.textContent = '';
                    }
                    tr.appendChild(tdGroup);
                    const tdNotes = document.createElement("td");
                    const notesBtn = document.createElement("button");
                    notesBtn.textContent = "备注";
                    notesBtn.onclick = () => {
                        const modal = new Modal(self.app);
                        modal.contentEl.style.padding = "24px";
                        modal.contentEl.createEl("h2", { text: v.word + ' 备注' });
                        const textarea = document.createElement('textarea');
                        textarea.style.width = '100%';
                        textarea.style.height = '120px';
                        textarea.value = v.notes || '';
                        modal.contentEl.appendChild(textarea);
                        const saveBtn = document.createElement('button');
                        saveBtn.textContent = '保存';
                        saveBtn.style.marginTop = '16px';
                        saveBtn.onclick = () => {
                            v.notes = textarea.value;
                            setTimeout(() => { self.plugin.saveVocabData(); }, 0);
                            modal.close();
                            renderTable();
                        };
                        modal.contentEl.appendChild(saveBtn);
                        modal.open();
                    };
                    tdNotes.appendChild(notesBtn);
                    tr.appendChild(tdNotes);
                    const opTd = document.createElement("td");
                    opTd.style.display = "flex";
                    opTd.style.gap = "6px";
                    opTd.style.flexDirection = "row";
                    opTd.style.alignItems = "center";
                    // 删除
                    const delBtn = document.createElement("button");
                    delBtn.textContent = "删除";
                    delBtn.onclick = () => {
                        const idx = self.plugin.vocabData.vocabBook.findIndex(
                            (item: any) => item.word === v.word && item.translation === v.translation
                        );
                        if (idx !== -1) {
                            if (!self.plugin.trashData.vocabBookTrash) self.plugin.trashData.vocabBookTrash = [];
                            self.plugin.trashData.vocabBookTrash.push({...v});
                            self.plugin.vocabData.vocabBook.splice(idx, 1);
                            renderTable();
                            setTimeout(() => { self.plugin.saveVocabData(); self.plugin.saveTrashData(); }, 0);
                        } else {
                            new Notice('删除失败：未找到该单词');
                        }
                    };
                    opTd.appendChild(delBtn);
                    // 编辑分组
                    const groupBtn = document.createElement("button");
                    groupBtn.textContent = "分组";
                    groupBtn.onclick = () => {
                        new VocabGroupOverviewModal(
                            self.app,
                            self.plugin,
                            v,
                            () => {
                                tdGroup.textContent = (Array.isArray(v.group) && v.group.length > 0) ? v.group.map((g: string[]) => g.join('/')).join(' | ') : '';
                                setTimeout(() => {
                                    self.plugin.saveVocabData();
                                    renderTable();
                                }, 0);
                            }
                        ).open();
                    };
                    opTd.appendChild(groupBtn);
                    // 标记掌握
                    const masteredBtn = document.createElement("button");
                    masteredBtn.textContent = v.mastered ? "已掌握" : "未掌握";
                    masteredBtn.style.background = v.mastered ? "#4caf50" : "#eee";
                    masteredBtn.onclick = () => {
                        v.mastered = !v.mastered;
                        masteredBtn.textContent = v.mastered ? "已掌握" : "未掌握";
                        masteredBtn.style.background = v.mastered ? "#4caf50" : "#eee";
                        setTimeout(() => {
                            self.plugin.saveVocabData();
                        }, 0);
                    };
                    opTd.appendChild(masteredBtn);
                    tr.appendChild(opTd);
            tbody.appendChild(tr);
        }
        table.appendChild(tbody);
            }

            container.onscroll = null;
            renderRows();
            container.appendChild(table);

            // 分页控件始终渲染在表格下方
            const pager = tableDiv.createEl("div");
            pager.style.margin = "12px 0";
            pager.style.textAlign = "center";
            pager.style.background = '#fff';
            pager.style.position = 'relative';
            pager.style.zIndex = '10';
            if (this._isMaximized) {
                pager.style.margin = '0 0 12px 0';
                pager.style.padding = '12px 0 12px 0';
                pager.style.boxShadow = '0 2px 8px #0001';
            } else {
                pager.style.margin = '12px 0';
                pager.style.padding = '';
                pager.style.boxShadow = '';
            }
            pager.innerHTML = `第 ${currentPage} / ${totalPages} 页`;
            // 页码跳转输入框
            const pageInput = document.createElement('input');
            pageInput.type = 'number';
            pageInput.min = '1';
            pageInput.max = String(totalPages);
            pageInput.value = String(currentPage);
            pageInput.style.width = '48px';
            pageInput.style.margin = '0 8px';
            pageInput.style.fontSize = '1em';
            pageInput.style.verticalAlign = 'middle';
            pageInput.title = '跳转到指定页';
            pageInput.onkeydown = (e) => {
                if (e.key === 'Enter') {
                    let val = parseInt(pageInput.value);
                    if (isNaN(val) || val < 1) val = 1;
                    if (val > totalPages) val = totalPages;
                    if (val !== currentPage) {
                        currentPage = val;
                        renderTable();
                    }
                }
            };
            pageInput.onblur = () => {
                let val = parseInt(pageInput.value);
                if (isNaN(val) || val < 1) val = 1;
                if (val > totalPages) val = totalPages;
                if (val !== currentPage) {
                    currentPage = val;
                    renderTable();
                }
            };
            pager.appendChild(pageInput);
            if (currentPage > 1) {
                const prevBtn = pager.createEl("button", { text: "上一页" });
                prevBtn.onclick = () => { currentPage--; renderTable(); };
            }
            if (currentPage < totalPages) {
                const nextBtn = pager.createEl("button", { text: "下一页" });
                nextBtn.onclick = () => { currentPage++; renderTable(); };
            }
            applyMaximizedStyle();
        };

        let searchMatches: number[] = [];
        let searchIndex = 0;
        let highlightActive = true;

        // 新增：全局变量 filteredData，渲染和分页都用它
        let filteredData: any[] = this.plugin && this.plugin.vocabData && Array.isArray(this.plugin.vocabData.vocabBook) ? this.plugin.vocabData.vocabBook : [];
        // performSearch 需在 modeBtn/searchMode 可见作用域下定义
        performSearch = () => {
            const searchTerm = searchInput.value.trim();
            const data = this.plugin.vocabData.vocabBook || [];
            if (!searchTerm) {
                filteredData = data;
                searchMatches = [];
                for (let i = 0; i < filteredData.length; i++) {
                    searchMatches.push(i);
                }
                searchIndex = 0;
                highlightActive = false;
                currentPage = 1; // 搜索后重置页码
                updateResultInfo();
                renderTable();
                return;
            }
            // --- 链式过滤实现 ---
            // 以 \\ 分割所有条件
            const filters = searchTerm.split('\\').map(s => s.trim()).filter(Boolean);
            let result = data;
            for (const filter of filters) {
                if (/^(\*+)$/.test(filter)) {
                    // 只 *、**、***，按单词数过滤
                    const starCount = filter.length;
                    result = result.filter((item: any) => {
                        const word = (item.word || '').replace(/\s+/g, ' ').trim();
                        return word && word.split(' ').length === starCount;
                    });
                } else if (/^\*[a-zA-Z]$/.test(filter)) {
                    // *a 这种，首字母过滤
                    const firstLetter = filter[1].toLowerCase();
                    result = result.filter((item: any) => {
                        const word = (item.word || '').trim();
                        return word && word[0].toLowerCase() === firstLetter;
                    });
                } else {
                    // 其它，区分模糊搜索和单词/短语名搜索
                    const lowerTerm = filter.toLowerCase();
                    if (searchMode === 'fuzzy') {
                        result = result.filter((item: any) => {
                            const word = (item.word || '').toLowerCase();
                            const translation = (item.translation || '').toLowerCase();
                            const example = (item.example || '').toLowerCase();
                            return word.includes(lowerTerm) || translation.includes(lowerTerm) || example.includes(lowerTerm);
                        });
                    } else {
                        // 只匹配单词名和短语名
                        result = result.filter((item: any) => {
                            const word = (item.word || '').toLowerCase();
                            const phrase = (item.phrase || '').toLowerCase();                       
                            return word.includes(lowerTerm) || phrase.includes(lowerTerm);
                        });
                    }
                }
            }
            filteredData = result;
            searchMatches = [];
            for (let i = 0; i < filteredData.length; i++) {
                searchMatches.push(i);
            }
            currentPage = 1;
            searchIndex = 0;
            highlightActive = true;
            if (searchMatches.length > 0) {
                const matchPage = Math.floor(searchIndex / pageSize) + 1;
                if (currentPage !== matchPage) {
                    currentPage = matchPage;
                }
            }
            updateResultInfo();
                renderTable();
        };
        // 现在 performSearch 已赋值，可以安全设置 modeBtn.onclick
        modeBtn.onclick = () => {
            if (searchMode === 'fuzzy') {
                searchMode = 'word';
                modeBtn.textContent = '单词名和短语名搜索';
            } else {
                searchMode = 'fuzzy';
                modeBtn.textContent = '模糊搜索';
            }
            performSearch();
        };
        const gotoMatch = (idxDelta: number) => {
            if (searchMatches.length === 0) return;
            searchIndex = (searchIndex + idxDelta + searchMatches.length) % searchMatches.length;
            // 自动跳转到包含当前 searchIndex 匹配项的页码
            const matchGlobalIndex = searchMatches[searchIndex];
            const matchPage = Math.floor(matchGlobalIndex / pageSize) + 1;
            if (currentPage !== matchPage) {
                currentPage = matchPage;
            }
            highlightActive = true;
            updateResultInfo();
            renderTable();
        };
        const updateResultInfo = () => {
            if (searchMatches.length > 0) {
                resultInfo.textContent = `${searchIndex + 1} of ${searchMatches.length}`;
            } else {
                resultInfo.textContent = '无结果';
            }
        };
        // 防止多次绑定事件
        upBtn.onclick = null;
        downBtn.onclick = null;
        searchBtn.onclick = null;
        searchInput.onkeydown = null;
        upBtn.onclick = () => gotoMatch(-1);
        downBtn.onclick = () => gotoMatch(1);
        searchBtn.onclick = performSearch;
        searchInput.onkeydown = (e: KeyboardEvent) => {
            if (e.key === 'Enter') {
                if (e.shiftKey) {
                    gotoMatch(-1); // Shift+Enter 上一个
                } else {
                    gotoMatch(1); // Enter 下一个
                }
            }
        };

        // 复习模式
        const startReview = (missOnly = false) => {
            let data = this.plugin.vocabData.vocabBook || [];
            if (missOnly) data = data.filter((v: any) => !v.mastered);
            if (!data.length) {
                new Notice(missOnly ? "没有未掌握词条" : "词汇本为空");
                return;
            }
            let idx = 0;
            const reviewModal = new Modal(this.app);
            const show = () => {
                reviewModal.contentEl.empty();
                const v = data[idx];
                reviewModal.contentEl.createEl("h3", { text: v.word });
                reviewModal.contentEl.createEl("div", { text: v.example || "" }).style.margin = "12px 0";
                const showAnsBtn = reviewModal.contentEl.createEl("button", { text: "显示释义" });
                showAnsBtn.onclick = () => {
                    reviewModal.contentEl.createEl("div", { text: v.translation, cls: "review-ans" });
                };
                const masteredBtn = reviewModal.contentEl.createEl("button", { text: v.mastered ? "已掌握" : "未掌握" });
                masteredBtn.onclick = async () => {
                    v.mastered = !v.mastered;
                    await this.plugin.saveVocabData();
                    show();
                };
                const prevBtn = reviewModal.contentEl.createEl("button", { text: "上一个" });
                prevBtn.onclick = () => { idx = (idx - 1 + data.length) % data.length; show(); };
                const nextBtn = reviewModal.contentEl.createEl("button", { text: "下一个" });
                nextBtn.onclick = () => { idx = (idx + 1) % data.length; show(); };
            };
            show();
            reviewModal.open();
        };

        // 事件绑定
        const reviewBtn = btnBar.createEl("button", { text: "复习模式" });
        reviewBtn.onclick = () => startReview();
        const reviewMissBtn = btnBar.createEl("button", { text: "查漏补缺" });
        reviewMissBtn.onclick = () => startReview(true);

        // 初始渲染
        renderTable();

        // 支持自由拖动
        modalEl.style.resize = "both";
        modalEl.style.overflow = "auto";
        modalEl.style.position = "absolute";
        modalEl.style.background = "#fff";
        modalEl.style.zIndex = String(Date.now());
        makeModalDraggable(modalEl);
        // 隐藏遮罩层
        const bg = modalEl.parentElement?.querySelector('.modal-bg') as HTMLElement;
        if (bg) bg.style.display = 'none';
        const h2 = header.querySelector("h2");
        if (h2) {
            h2.style.cursor = "move";
        }
        // 全部分组按钮
        const allGroupsBtn = btnBar.createEl("button", { text: "全部分组" });
        allGroupsBtn.onclick = () => {
            new GroupBrowserModal(this.app, this.plugin).open();
        };

        // === 拖动标志插入 start ===
        let dragHandle = modalEl.querySelector('.modal-drag-handle') as HTMLElement;
        if (!dragHandle) {
            dragHandle = document.createElement('div');
            dragHandle.className = 'modal-drag-handle';
            dragHandle.textContent = '≡';
            dragHandle.style.position = 'absolute';
            dragHandle.style.left = '16px';
            dragHandle.style.top = '16px';
            dragHandle.style.fontSize = '1.5em';
            dragHandle.style.cursor = 'grab';
            dragHandle.style.userSelect = 'none';
            dragHandle.style.zIndex = '100001';
            modalEl.appendChild(dragHandle);
        }
        makeModalDraggable(modalEl, dragHandle);
        // === 拖动标志插入 end ===

        // 把 modeBtn 插入到 searchBoxWrapper 最前面
        searchBoxWrapper.insertBefore(modeBtn, searchBoxWrapper.firstChild);

        // === 强制最大化样式（彻底全屏） ===
        if (this._isMaximized) {
            modalEl.style.position = 'fixed';
            modalEl.style.left = '0';
            modalEl.style.top = '0';
            modalEl.style.width = '100vw';
            modalEl.style.height = '100vh';
            modalEl.style.maxWidth = '100vw';
            modalEl.style.maxHeight = '100vh';
            modalEl.style.borderRadius = '0';
            modalEl.style.zIndex = '999999';
            contentEl.style.width = '100vw';
            contentEl.style.height = '100vh';
            contentEl.style.minWidth = '0';
            contentEl.style.minHeight = '0';
            contentEl.style.maxWidth = '100vw';
            contentEl.style.maxHeight = '100vh';
            contentEl.style.overflow = 'hidden';
        }
    }
    onClose() {
        if (globalVocabBookModal === this) globalVocabBookModal = null;
        this.contentEl.empty();
    }
    onClickOutside() {
        // 阻止点击遮罩关闭
    }
}

function exportVocabBook(plugin: any) {
    const data = JSON.stringify(plugin.vocabData.vocabBook || [], null, 2);
    const blob = new Blob([data], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'vocabBook.json';
    a.click();
    URL.revokeObjectURL(url);
}

async function importVocabBook(plugin: any, file: File) {
    const text = await file.text();
    let arr = [];
    try { arr = JSON.parse(text); } catch {}
    if (!Array.isArray(arr)) { new Notice('导入格式错误'); return; }
    // 自动字段映射，兼容任意结构
    const mapped = arr.map((item: any) => {
        // 标准字段
        const word = item.word || item.headWord || item.front || item.term || item.vocab || "";
        const translation = item.translation || item.meaning || item.content || item.back || item.definition ||
            (Array.isArray(item.translations) ? item.translations.map((t: any) => (t.type ? `[${t.type}]` : "") + t.translation).join('; ') : "");
        const example = item.example || item.sentence || item.eg || "";
        let group = item.group || item.tag || item.category || "";
        // 修正 group 字段为分级数组
        if (typeof group === 'string' && group.includes('/')) {
            group = group.split('/').map((s: string) => s.trim()).filter(Boolean);
        } else if (Array.isArray(group) && group.length === 1 && typeof group[0] === 'string' && group[0].includes('/')) {
            group = group[0].split('/').map((s: string) => s.trim()).filter(Boolean);
        }
        const mastered = item.mastered || item.isMastered || false;
        // 友好拼接 phrases
        const phrases = Array.isArray(item.phrases) ? item.phrases.map((p: any) => `${p.phrase || ''}: ${p.translation || ''}`).join('; ') : '';
        // 收集所有未标准化字段
        const stdKeys = new Set(['word','headWord','front','term','vocab','translation','meaning','content','back','definition','translations','example','sentence','eg','group','tag','category','mastered','isMastered','phrases','notes']);
        let extraNotes = [];
        for (const k in item) {
            if (!stdKeys.has(k)) {
                let v = item[k];
                if (typeof v === 'object') {
                    try { v = JSON.stringify(v); } catch {} 
                }
                extraNotes.push(`${k}: ${v}`);
            }
        }
        // 合并 notes 字段
        let notes = item.notes || '';
        if (phrases) notes += (notes ? '\n' : '') + '短语: ' + phrases;
        if (extraNotes.length) notes += (notes ? '\n' : '') + extraNotes.join('\n');
        return {
            word,
            translation,
            example,
            group,
            addedAt: Date.now(),
            notes,
            mastered
        };
    });
    plugin.vocabData.vocabBook = mapped;
    await plugin.saveVocabData();
    new Notice('导入成功');
}

async function addToVocabBook(plugin: any, item: any) {
    if (!plugin.vocabData.vocabBook) plugin.vocabData.vocabBook = [];
    if (!plugin.vocabData.vocabBook.find((v: any) => v.word === item.word && v.translation === item.translation)) {
        plugin.vocabData.vocabBook.push(item);
        await plugin.saveVocabData();
        new Notice('已加入词汇本');
    } else {
        new Notice('该词条已存在词汇本');
    }
}

class GroupSelectModal extends Modal {
    app: App;
    allGroups: string[][]; // 所有分组路径
    maxDepth: number = 5;
    onSelect: (groupPath: string[]) => void;
    currentPath: string[] = [];
    _isMinimized: boolean = false;
    _isMaximized: boolean = false;
    _modalId: string;
    onReset?: () => void; // 新增
    constructor(app: App, allGroups: string[][], onSelect: (groupPath: string[]) => void, startPath: string[] = [], onReset?: () => void) {
        super(app);
        this.app = app;
        // 保证 allGroups 是 string[][]，每个元素为一条分组路径（如 ["A", "B"]），避免三维嵌套
        const newAllGroups: string[][] = [];
        for (const item of allGroups) {
            if (Array.isArray(item)) {
                // 兼容历史数据：item 可能是 string[] 或 string[][]
                if (item.length > 0 && Array.isArray(item[0])) {
                    // string[][]
                    for (const g of item as any[]) {
                        if (Array.isArray(g) && g.length > 0 && g.every((x: any) => typeof x === 'string')) {
                            newAllGroups.push(g as string[]);
                        }
                    }
                } else if (item.length > 0 && typeof item[0] === 'string') {
                    // string[]
                    newAllGroups.push(item as string[]);
                }
            } else if (typeof item === 'string' && item) {
                // 兼容历史数据：单条字符串
                newAllGroups.push((item as string).split('/').map((s: string) => s.trim()).filter(Boolean));
            }
        }
        this.allGroups = newAllGroups;
        this.onSelect = onSelect;
        this.currentPath = [...startPath];
        this.onReset = onReset;
        this._modalId = getUniqueModalId();
    }
    onOpen() {
        this.renderLayer();
        // 拖动
        makeModalDraggable(this.modalEl);
        // 最小化/最大化按钮
        const header = this.contentEl.querySelector('h2') || this.contentEl.firstElementChild;
        if (header) {
            const btnGroup = document.createElement('div');
            btnGroup.style.display = 'flex';
            btnGroup.style.gap = '8px';
            btnGroup.style.position = 'absolute';
            btnGroup.style.top = '16px';
            btnGroup.style.right = '32px';
            // 最小化
            const minBtn = document.createElement('button');
            minBtn.textContent = '–';
            minBtn.title = '最小化';
            minBtn.onclick = () => {
                this._isMinimized = true;
                this.modalEl.style.display = 'none';
                let bar = document.getElementById('youdao-modal-taskbar');
                if (!bar) {
                    bar = document.createElement('div');
                    bar.id = 'youdao-modal-taskbar';
                    bar.style.position = 'fixed';
                    bar.style.left = '50%';
                    bar.style.transform = 'translateX(-50%)';
                    bar.style.bottom = '0';
                    bar.style.height = '44px';
                    bar.style.background = 'rgba(255,255,255,0.95)';
                    bar.style.zIndex = '99999';
                    bar.style.display = 'flex';
                    bar.style.alignItems = 'center';
                    bar.style.gap = '12px';
                    bar.style.padding = '0 16px';
                    document.body.appendChild(bar);
                }
                const restoreBtn = document.createElement('button');
                restoreBtn.textContent = '分组选择';
                restoreBtn.onclick = () => {
                    this.modalEl.style.display = '';
                    this._isMinimized = false;
                    restoreBtn.remove();
                };
                bar.appendChild(restoreBtn);
            };
            // 最大化
            const maxBtn = document.createElement('button');
            maxBtn.textContent = '☐';
            maxBtn.title = '最大化';
            maxBtn.onclick = () => {
                this._isMaximized = !this._isMaximized;
                if (this._isMaximized) {
                    this.modalEl.style.width = '98vw';
                    this.modalEl.style.height = '96vh';
                    this.modalEl.style.left = '1vw';
                    this.modalEl.style.top = '2vh';
                    maxBtn.style.fontWeight = 'bold';
                } else {
                    this.modalEl.style.width = '';
                    this.modalEl.style.height = '';
                    this.modalEl.style.left = '';
                    this.modalEl.style.top = '';
                    maxBtn.style.fontWeight = '';
                }
            };
            btnGroup.appendChild(minBtn);
            btnGroup.appendChild(maxBtn);
            this.contentEl.appendChild(btnGroup);
        }
        // 禁止遮罩关闭
        const bg = this.modalEl.parentElement?.querySelector('.modal-bg') as HTMLElement;
        if (bg) bg.style.pointerEvents = 'none';
        // 在底部添加重置分组按钮
        if (typeof this.onReset === 'function') {
            const resetDiv = this.contentEl.createEl('div');
            resetDiv.style.marginTop = '24px';
            resetDiv.style.textAlign = 'center';
            const resetBtn = resetDiv.createEl('button', { text: '重置分组（清空）' });
            resetBtn.style.background = '#e53935';
            resetBtn.style.color = '#fff';
            resetBtn.style.padding = '6px 18px';
            resetBtn.style.borderRadius = '6px';
            resetBtn.style.fontWeight = 'bold';
            resetBtn.onclick = () => {
                this.onReset && this.onReset();
                this.close();
            };
        }
    }
    renderLayer() {
        const { contentEl } = this;
        contentEl.empty();
        contentEl.style.padding = "24px";
        contentEl.style.background = "#f8fafc";
        contentEl.style.minWidth = "320px";
        contentEl.style.maxWidth = "90vw";
        contentEl.style.maxHeight = "80vh";
        contentEl.style.overflowY = "auto";
        // 返回按钮
        if (this.currentPath.length === 0) {
            // 第一层，返回分组管理界面
            const backBtn = contentEl.createEl("button", { text: "← 返回" });
            backBtn.style.marginBottom = "12px";
            backBtn.onclick = () => {
                this.close();
                if ((window as any).__lastVocabGroupOverviewModal) {
                    (window as any).__lastVocabGroupOverviewModal();
                }
            };
        } else {
            // 其它层，返回上一级
            const backBtn = contentEl.createEl("button", { text: "← 返回" });
            backBtn.style.marginBottom = "12px";
            backBtn.onclick = () => {
                this.currentPath.pop();
                this.renderLayer();
            };
        }
        // 当前路径显示
        if (this.currentPath.length > 0) {
            const pathDiv = contentEl.createEl("div", { text: "当前位置：" + this.currentPath.join(" / ") });
            pathDiv.style.marginBottom = "8px";
            pathDiv.style.fontWeight = "bold";
        }
        // 输入框：新建/输入本层分组
        const inputDiv = contentEl.createEl("div");
        inputDiv.style.display = "flex";
        inputDiv.style.gap = "8px";
        const input = inputDiv.createEl("input");
        input.type = "text";
        input.placeholder = `输入${this.currentPath.length + 1}级分组名（最多5级）`;
        input.style.flex = "1";
        const okBtn = inputDiv.createEl("button", { text: this.currentPath.length === this.maxDepth - 1 ? "确定" : "进入/新建" });
        okBtn.onclick = () => {
            const val = input.value.trim();
            if (!val) return;
            if (this.currentPath.length < this.maxDepth - 1) {
                this.currentPath.push(val);
                this.renderLayer();
            } else {
                // 最后一层，直接确定
                this.onSelect([...this.currentPath, val]);
                this.close();
            }
        };
        // 回车快捷键
        input.onkeydown = (e: KeyboardEvent) => {
            if (e.key === 'Enter') okBtn.click();
        };
        contentEl.appendChild(inputDiv);
        // 下拉：本层已有分组
        const siblings = this.getSiblings(this.currentPath);
        if (siblings.length > 0) {
            const selDiv = contentEl.createEl("div");
            selDiv.style.marginTop = "12px";
            selDiv.createEl("div", { text: "选择已有分组：" });
            siblings.forEach((name) => {
                // 保证 name 一定为字符串
                if (Array.isArray(name)) name = name.join('/');
                name = String(name);
                console.log('[siblings.forEach] name:', name, 'typeof:', typeof name, 'isArray:', Array.isArray(name));
                const btn = selDiv.createEl("button", { text: name });
                btn.style.margin = "4px 8px 4px 0";
                btn.onclick = () => {
                        this.currentPath.push(name);
                        this.renderLayer();
                };
            });
        }
        // 新增：每层都加确定按钮（第一级也能用）
        const confirmDiv = contentEl.createEl("div");
        confirmDiv.style.marginTop = "18px";
        const confirmBtn = confirmDiv.createEl("button", { text: "确定（使用当前路径）" });
        confirmBtn.onclick = () => {
            if (this.currentPath.length > 0) {
                this.onSelect([...this.currentPath]);
                this.close();
            }
        };
        confirmBtn.disabled = this.currentPath.length === 0;
    }
    // 获取本层已有分组名
    // 获取本层已有分组名（允许同名，严格分层）
    getSiblings(path: string[]): string[] {
        const siblings: string[] = [];
        for (const g of this.allGroups) {
            if (g.length > path.length && g.slice(0, path.length).join("\u0000") === path.join("\u0000")) {
                console.log('[getSiblings] g:', g, 'path:', path, 'g[path.length]:', g[path.length]);
                siblings.push(g[path.length]);
            }
        }
        return Array.from(new Set(siblings.map(s => Array.isArray(s) ? s.join('/') : s)));
    }
}

// 新增：分组概览弹窗
class VocabGroupOverviewModal extends Modal {
    plugin: any;
    vocabItem: any;
    onUpdate: () => void;
    _isMinimized: boolean = false;
    _isMaximized: boolean = false;
    _modalId: string;
    constructor(app: App, plugin: any, vocabItem: any, onUpdate: () => void) {
        super(app);
        this.plugin = plugin;
        this.vocabItem = vocabItem;
        this.onUpdate = onUpdate;
        this._modalId = getUniqueModalId();
    }
    onOpen() {
        const { contentEl, modalEl } = this;
        contentEl.empty();
        contentEl.style.padding = "24px";
        contentEl.style.background = "#f8fafc";
        contentEl.style.minWidth = "400px";
        contentEl.style.maxWidth = "90vw";
        contentEl.style.maxHeight = "80vh";
        contentEl.style.overflowY = "auto";
        // 标题栏和按钮组
        const header = contentEl.createEl("div");
        header.style.display = "flex";
        header.style.justifyContent = "space-between";
        header.style.alignItems = "center";
        header.style.marginBottom = "16px";
        const title = header.createEl("h2", { text: "分组管理" });
        title.style.margin = '0';
        // 按钮组
        const btnGroup = header.createEl("div");
        btnGroup.style.display = "flex";
        btnGroup.style.gap = "8px";
        // 最小化按钮
        const minBtn = btnGroup.createEl("button", { text: "–" });
        minBtn.title = "最小化";
        minBtn.onclick = () => {
            this._isMinimized = true;
            modalEl.style.display = 'none';
            let bar = document.getElementById('youdao-modal-taskbar');
            if (!bar) {
                bar = document.createElement('div');
                bar.id = 'youdao-modal-taskbar';
                bar.style.position = 'fixed';
                bar.style.left = '50%';
                bar.style.transform = 'translateX(-50%)';
                bar.style.bottom = '0';
                bar.style.height = '44px';
                bar.style.background = 'rgba(255,255,255,0.95)';
                bar.style.zIndex = '99999';
                bar.style.display = 'flex';
                bar.style.alignItems = 'center';
                bar.style.gap = '12px';
                bar.style.padding = '0 16px';
                document.body.appendChild(bar);
            }
            const restoreBtn = document.createElement('button');
            restoreBtn.textContent = '分组管理';
            restoreBtn.onclick = () => {
                modalEl.style.display = '';
                this._isMinimized = false;
                restoreBtn.remove();
            };
            bar.appendChild(restoreBtn);
        };
        // 最大化按钮
        const maxBtn = btnGroup.createEl("button", { text: "☐" });
        maxBtn.title = "最大化";
        maxBtn.onclick = () => {
            this._isMaximized = !this._isMaximized;
            if (this._isMaximized) {
                modalEl.style.width = '98vw';
                modalEl.style.height = '96vh';
                modalEl.style.left = '1vw';
                modalEl.style.top = '2vh';
                maxBtn.style.fontWeight = 'bold';
            } else {
                modalEl.style.width = '';
                modalEl.style.height = '';
                modalEl.style.left = '';
                modalEl.style.top = '';
                maxBtn.style.fontWeight = '';
            }
        };
        // 拖动支持
        makeModalDraggable(modalEl, title);
        // 分组路径列表
        let groupArr: string[][] = [];
        if (Array.isArray(this.vocabItem.group)) {
            groupArr = this.vocabItem.group;
        } else if (typeof this.vocabItem.group === 'string' && this.vocabItem.group) {
            groupArr = this.vocabItem.group
                .split('\n')
                .map((line: any) => (line as string).split('/').map((s: any) => (s as string).trim()).filter((s: any) => Boolean(s)))
                .filter((arr: any) => arr.length > 0);
        }
        if (!groupArr.length) {
            const info = contentEl.createEl("div", { text: "暂无分组路径，可新建分组。" });
            info.style.marginBottom = "16px";
            const createBtn = contentEl.createEl("button", { text: "创建分组" });
            createBtn.style.background = "#1a73e8";
            createBtn.style.color = "#fff";
            createBtn.style.padding = "8px 24px";
            createBtn.style.fontSize = "1.1em";
            createBtn.style.border = "none";
            createBtn.style.borderRadius = "6px";
            createBtn.style.cursor = "pointer";
            createBtn.onclick = () => {
                this.close();
                (window as any).__lastVocabGroupOverviewModal = () => {
                    new VocabGroupOverviewModal(this.app, this.plugin, this.vocabItem, this.onUpdate).open();
                };
                new GroupSelectModal(
                    this.app,
                    this.plugin.vocabData.vocabBook.map((item: any) => Array.isArray(item.group) ? item.group : (typeof item.group === 'string' && item.group ? item.group.split('/').map((s: string) => s.trim()).filter(Boolean) : [])),
                    (selectedPath) => {
                        console.log('[分组新建] typeof group:', typeof this.vocabItem.group, 
                            'isArray:', Array.isArray(this.vocabItem.group), 
                            'group:', this.vocabItem.group, 
                            'selectedPath:', selectedPath, 
                            'stack:', new Error().stack);
                        if (!Array.isArray(this.vocabItem.group)) this.vocabItem.group = [];
                        this.vocabItem.group.push([...selectedPath]);
                        console.log('[分组新建后] typeof group:', typeof this.vocabItem.group, 
                            'isArray:', Array.isArray(this.vocabItem.group), 
                            'group:', this.vocabItem.group, 
                            'stack:', new Error().stack);
                        setTimeout(() => { this.plugin.saveVocabData(); this.onUpdate(); }, 0);
                    },
                    [],
                    undefined
                ).open();
            };
        } else {
            const listDiv = contentEl.createEl("div");
            listDiv.style.marginBottom = "16px";
            listDiv.createEl("div", { text: "已有分组路径：" });
            groupArr.forEach((path, idx) => {
                const row = listDiv.createEl("div");
                row.style.display = "flex";
                row.style.alignItems = "center";
                row.style.margin = "8px 0";
                const text = Array.isArray(path) ? path.join('/') : String(path);
                row.createEl("span", { text });
                // 修改分组按钮
                const editBtn = row.createEl("button", { text: "修改分组" });
                editBtn.style.marginLeft = "12px";
                editBtn.style.background = "#eee";
                editBtn.style.border = "none";
                editBtn.style.borderRadius = "4px";
                editBtn.style.cursor = "pointer";
                editBtn.onclick = () => {
                    this.close();
                    (window as any).__lastVocabGroupOverviewModal = () => {
                        new VocabGroupOverviewModal(this.app, this.plugin, this.vocabItem, this.onUpdate).open();
                    };
                    new GroupSelectModal(
                        this.app,
                        this.plugin.vocabData.vocabBook.map((item: any) => Array.isArray(item.group) ? item.group : (typeof item.group === 'string' && item.group ? item.group.split('/').map((s: string) => s.trim()).filter(Boolean) : [])),
                        (selectedPath) => {
                            console.log('[分组修改] typeof group:', typeof this.vocabItem.group, 
                                'isArray:', Array.isArray(this.vocabItem.group), 
                                'group:', this.vocabItem.group, 
                                'selectedPath:', selectedPath, 
                                'idx:', idx, 
                                'stack:', new Error().stack);
                            if (!Array.isArray(this.vocabItem.group)) this.vocabItem.group = [];
                            this.vocabItem.group[idx] = [...selectedPath];
                            console.log('[分组修改后] typeof group:', typeof this.vocabItem.group, 
                                'isArray:', Array.isArray(this.vocabItem.group), 
                                'group:', this.vocabItem.group, 
                                'stack:', new Error().stack);
                            setTimeout(() => { this.plugin.saveVocabData(); this.onUpdate(); }, 0);
                        },
                        path,
                        undefined
                    ).open();
                };
                // 删除分组按钮
                const delBtn = row.createEl("button", { text: "删除分组" });
                delBtn.style.marginLeft = "8px";
                delBtn.style.background = "#eee";
                delBtn.style.border = "none";
                delBtn.style.borderRadius = "4px";
                delBtn.style.cursor = "pointer";
                delBtn.onclick = () => {
                    if (Array.isArray(this.vocabItem.group)) {
                        this.vocabItem.group.splice(idx, 1);
                        setTimeout(() => { this.plugin.saveVocabData(); this.onUpdate(); }, 0);
                    }
                };
            });
            const createBtn = contentEl.createEl("button", { text: "创建分组" });
            createBtn.style.background = "#1a73e8";
            createBtn.style.color = "#fff";
            createBtn.style.padding = "8px 24px";
            createBtn.style.fontSize = "1.1em";
            createBtn.style.border = "none";
            createBtn.style.borderRadius = "6px";
            createBtn.style.cursor = "pointer";
            createBtn.style.marginTop = "16px";
            createBtn.onclick = () => {
                this.close();
                (window as any).__lastVocabGroupOverviewModal = () => {
                    new VocabGroupOverviewModal(this.app, this.plugin, this.vocabItem, this.onUpdate).open();
                };
                new GroupSelectModal(
                    this.app,
                    this.plugin.vocabData.vocabBook.map((item: any) => Array.isArray(item.group) ? item.group : (typeof item.group === 'string' && item.group ? item.group.split('/').map((s: string) => s.trim()).filter(Boolean) : [])),
                    (selectedPath) => {
                        if (!Array.isArray(this.vocabItem.group)) this.vocabItem.group = [];
                        this.vocabItem.group.push(selectedPath);
                        setTimeout(() => { this.plugin.saveVocabData(); this.onUpdate(); }, 0);
                    },
                    [],
                    undefined
                ).open();
            };
        }
        // 拖动支持
        makeModalDraggable(modalEl, title);
    }
}

// 新增：分组逐层浏览弹窗
class GroupBrowserModal extends Modal {
    plugin: any;
    currentPath: string[] = [];
    maxDepth: number = 5;
    _isMinimized: boolean = false;
    _isMaximized: boolean = false;
    _modalId: string;
    constructor(app: App, plugin: any, startPath: string[] = []) {
        super(app);
        this.plugin = plugin;
        this.currentPath = [...startPath];
        this._modalId = getUniqueModalId();
    }
    onOpen() {
        this.renderLayer();
        makeModalDraggable(this.modalEl);
        // 最小化/最大化按钮
        const header = this.contentEl.querySelector('h2') || this.contentEl.firstElementChild;
        if (header) {
            const btnGroup = document.createElement('div');
            btnGroup.style.display = 'flex';
            btnGroup.style.gap = '8px';
            btnGroup.style.position = 'absolute';
            btnGroup.style.top = '16px';
            btnGroup.style.right = '32px';
            // 最小化
            const minBtn = document.createElement('button');
            minBtn.textContent = '–';
            minBtn.title = '最小化';
            minBtn.onclick = () => {
                this._isMinimized = true;
                this.modalEl.style.display = 'none';
                let bar = document.getElementById('youdao-modal-taskbar');
                if (!bar) {
                    bar = document.createElement('div');
                    bar.id = 'youdao-modal-taskbar';
                    bar.style.position = 'fixed';
                    bar.style.left = '50%';
                    bar.style.transform = 'translateX(-50%)';
                    bar.style.bottom = '0';
                    bar.style.height = '44px';
                    bar.style.background = 'rgba(255,255,255,0.95)';
                    bar.style.zIndex = '99999';
                    bar.style.display = 'flex';
                    bar.style.alignItems = 'center';
                    bar.style.gap = '12px';
                    bar.style.padding = '0 16px';
                    document.body.appendChild(bar);
                }
                const restoreBtn = document.createElement('button');
                restoreBtn.textContent = '分组浏览';
                restoreBtn.onclick = () => {
                    this.modalEl.style.display = '';
                    this._isMinimized = false;
                    restoreBtn.remove();
                };
                bar.appendChild(restoreBtn);
            };
            // 最大化
            const maxBtn = document.createElement('button');
            maxBtn.textContent = '☐';
            maxBtn.title = '最大化';
            maxBtn.onclick = () => {
                this._isMaximized = !this._isMaximized;
                if (this._isMaximized) {
                    this.modalEl.style.width = '98vw';
                    this.modalEl.style.height = '96vh';
                    this.modalEl.style.left = '1vw';
                    this.modalEl.style.top = '2vh';
                    maxBtn.style.fontWeight = 'bold';
                } else {
                    this.modalEl.style.width = '';
                    this.modalEl.style.height = '';
                    this.modalEl.style.left = '';
                    this.modalEl.style.top = '';
                    maxBtn.style.fontWeight = '';
                }
            };
            btnGroup.appendChild(minBtn);
            btnGroup.appendChild(maxBtn);
            this.contentEl.appendChild(btnGroup);
        }
        // 禁止遮罩关闭
        const bg = this.modalEl.parentElement?.querySelector('.modal-bg') as HTMLElement;
        if (bg) bg.style.pointerEvents = 'none';
    }
    renderLayer() {
        const { contentEl } = this;
        contentEl.empty();
        contentEl.style.padding = "24px";
        contentEl.style.background = "#f8fafc";
        contentEl.style.minWidth = "400px";
        contentEl.style.maxWidth = "90vw";
        contentEl.style.maxHeight = "80vh";
        contentEl.style.overflowY = "auto";
        // 返回按钮
        if (this.currentPath.length > 0) {
            const backBtn = contentEl.createEl("button", { text: "← 返回" });
            backBtn.style.marginBottom = "12px";
            backBtn.onclick = () => {
                this.currentPath.pop();
                this.renderLayer();
            };
        }
        // 路径显示
        if (this.currentPath.length > 0) {
            const pathDiv = contentEl.createEl("div", { text: "当前位置：" + this.currentPath.join(" / ") });
            pathDiv.style.marginBottom = "8px";
            pathDiv.style.fontWeight = "bold";
        }
        // 构建分组树
        const vocabBook = this.plugin.vocabData.vocabBook || [];
        // 收集所有分组路径
        let allGroups: string[][] = [];
        for (const v of vocabBook) {
            if (Array.isArray(v.group)) {
                // 兼容历史数据：v.group 可能是 string[] 或 string[][]
                if (v.group.length > 0 && Array.isArray(v.group[0])) {
                    // string[][]
                    for (const g of v.group as any[]) {
                        if (Array.isArray(g) && g.length > 0 && g.every((x: any) => typeof x === 'string')) {
                            allGroups.push(g as string[]);
                        }
                    }
                } else if (v.group.length > 0 && typeof v.group[0] === 'string') {
                    // string[]
                    allGroups.push(v.group as string[]);
                }
            } else if (typeof v.group === 'string' && v.group) {
                // 兼容历史数据：单条字符串
                allGroups.push((v.group as string).split('/').map((s: string) => s.trim()).filter(Boolean));
            }
        }
        // 保证 allGroups 是 string[][]，每个元素为一条分组路径（如 ["A", "B"]），避免三维嵌套
        allGroups = allGroups.map(g => {
            if (Array.isArray(g) && g.length === 1 && Array.isArray(g[0])) {
                return g[0];
            }
            return g;
        }).filter(g => Array.isArray(g) && g.length > 0 && g.every(x => typeof x === 'string'));
        console.log('[分组调试] 当前路径 currentPath:', this.currentPath, 'allGroups:', allGroups);
        // 获取本层所有分组名
        let siblings = (() => {
            const set = new Set<string>();
            for (const g of allGroups) {
                if (g.length > this.currentPath.length) {
                    // 只要前缀匹配即可
                    let match = true;
                    for (let i = 0; i < this.currentPath.length; i++) {
                        if (g[i] !== this.currentPath[i]) {
                            match = false;
                            break;
                        }
                    }
                    if (match) {
                        let name = g[this.currentPath.length];
                        if (Array.isArray(name)) name = name.join('/');
                        name = String(name);
                        set.add(name);
                    }
                }
            }
            const arr = Array.from(set);
            console.log('[分组调试] siblings 结果:', arr, '当前路径:', this.currentPath);
            return arr;
        })();
        // 新增：每一级都加分组搜索框
        let groupSearch = "";
        let filteredSiblings = siblings;
        if (siblings.length > 0 && this.currentPath.length < this.maxDepth) {
            const searchDiv = contentEl.createEl("div");
            searchDiv.style.marginBottom = "8px";
            const searchInput = searchDiv.createEl("input");
            searchInput.type = "text";
            searchInput.placeholder = "搜索分组名";
            searchInput.style.padding = "4px 8px";
            searchInput.style.fontSize = "1em";
            searchInput.oninput = () => {
                groupSearch = searchInput.value.trim().toLowerCase();
                filteredSiblings = siblings.filter(name => !groupSearch || name.toLowerCase().includes(groupSearch));
                console.log('[分组调试] 搜索后 filteredSiblings:', filteredSiblings, 'groupSearch:', groupSearch);
                // 重新渲染 siblings 按钮区域
                // 先移除旧的 siblings 按钮区域
                const oldBtns = Array.from(contentEl.querySelectorAll('.group-sibling-btn'));
                oldBtns.forEach(el => el.remove());
                // 重新插入过滤后的 siblings
                if (filteredSiblings.length > 0) {
                    filteredSiblings.forEach(name => {
                        if (Array.isArray(name)) name = name.join('/');
                        name = String(name);
                        console.log('[分组调试] 渲染按钮:', name, '当前路径:', this.currentPath);
                        const btn = contentEl.createEl("button", { text: name });
                        btn.classList.add('group-sibling-btn');
                        btn.style.margin = "4px 8px 4px 0";
                        btn.onclick = () => {
                            this.currentPath.push(name);
                            this.renderLayer();
                        };
                    });
                }
            };
        }
        // 选择已有分组
        if (siblings.length > 0 && this.currentPath.length < this.maxDepth) {
            contentEl.createEl("div", { text: "选择分组：" });
            filteredSiblings.forEach(name => {
                if (Array.isArray(name)) name = name.join('/');
                name = String(name);
                console.log('[分组调试] 渲染按钮:', name, '当前路径:', this.currentPath);
                const btn = contentEl.createEl("button", { text: name });
                btn.classList.add('group-sibling-btn');
                btn.style.margin = "4px 8px 4px 0";
                btn.onclick = () => {
                    this.currentPath.push(name);
                    this.renderLayer();
                };
            });
        } else {
            // 没有下一级，显示该分组链下所有单词
            let arr = vocabBook.filter((v: any) => {
                let groupPaths: string[][] = [];
                if (Array.isArray(v.group)) {
                    if (v.group.length > 0 && Array.isArray(v.group[0])) {
                        // string[][]
                        groupPaths = v.group as string[][];
                    } else if (v.group.length > 0 && typeof v.group[0] === 'string') {
                        // string[]
                        groupPaths = [v.group as string[]];
                    }
                } else if (typeof v.group === 'string' && v.group) {
                    groupPaths = [v.group.split('/').map((s: string) => s.trim()).filter(Boolean)];
                }
                // 只要有一条路径完全匹配 currentPath
                return groupPaths.some(g => 
                    g.length === this.currentPath.length && 
                    g.every((val, idx) => val === this.currentPath[idx])
                );
            });
            console.log('[分组调试] 末级分组，当前路径:', this.currentPath, '单词列表:', arr);
            // 新增：单词搜索框
            let wordSearch = "";
            let filteredArr = arr;
            const wordSearchDiv = contentEl.createEl("div");
            wordSearchDiv.style.marginBottom = "8px";
            const wordSearchInput = wordSearchDiv.createEl("input");
            wordSearchInput.type = "text";
            wordSearchInput.placeholder = "搜索单词";
            wordSearchInput.style.padding = "4px 8px";
            wordSearchInput.style.fontSize = "1em";
            wordSearchInput.oninput = () => {
                wordSearch = wordSearchInput.value.trim().toLowerCase();
                filteredArr = arr.filter((v: any) => v.word && v.word.toLowerCase().includes(wordSearch));
                console.log('[分组调试] 单词搜索 filteredArr:', filteredArr, 'wordSearch:', wordSearch);
                // 重新渲染单词列表区域
                const oldList = contentEl.querySelector('.group-word-list');
                if (oldList) oldList.remove();
                if (filteredArr.length > 0) {
                    const ul = contentEl.createEl("ul");
                    ul.classList.add('group-word-list');
                    ul.style.margin = "8px 0 8px 24px";
                    filteredArr.forEach((v: any) => {
                        const li = ul.createEl("li");
                        li.textContent = v.word + (v.translation ? `：${v.translation}` : "");
                    });
                } else {
                    const emptyDiv = contentEl.createEl("div", { text: "该分组下暂无单词" });
                    emptyDiv.classList.add('group-word-list');
                }
            };
            if (filteredArr.length > 0) {
                contentEl.createEl("div", { text: "单词列表：" });
                const ul = contentEl.createEl("ul");
                ul.classList.add('group-word-list');
                ul.style.margin = "8px 0 8px 24px";
                filteredArr.forEach((v: any) => {
                    const li = ul.createEl("li");
                    li.textContent = v.word + (v.translation ? `：${v.translation}` : "");
                });
            } else {
                const emptyDiv = contentEl.createEl("div", { text: "该分组下暂无单词" });
                emptyDiv.classList.add('group-word-list');
            }
        }
    }
}

// 在文件末尾添加样式
const vocabHighlightStyle = document.createElement('style');
vocabHighlightStyle.innerHTML = `.vocab-row-highlight { background: #ffe082 !important; }`;
document.head.appendChild(vocabHighlightStyle);

class YoudaoTranslateConfigModal extends Modal {
    plugin: any;
    constructor(app: App, plugin: any) {
        super(app);
        this.plugin = plugin;
    }
    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        contentEl.createEl('h2', { text: '有道翻译设置' });
        // 文本 appKey
        new Setting(contentEl)
            .setName('文本 appKey')
            .setDesc('你的有道文本翻译应用ID')
            .addText(text => text
                .setPlaceholder('appKey')
                .setValue(this.plugin.settings.textAppKey)
                .onChange(async (value) => {
                    this.plugin.settings.textAppKey = value;
                    await this.plugin.saveSettings();
                }));
        // 文本 appSecret
        new Setting(contentEl)
            .setName('文本 appSecret')
            .setDesc('你的有道文本翻译应用密钥')
            .addText(text => text
                .setPlaceholder('appSecret')
                .setValue(this.plugin.settings.textAppSecret)
                .onChange(async (value) => {
                    this.plugin.settings.textAppSecret = value;
                    await this.plugin.saveSettings();
                }));
        // 图片 appKey
        new Setting(contentEl)
            .setName('图片 appKey')
            .setDesc('你的有道图片翻译应用ID')
            .addText(text => text
                .setPlaceholder('imageAppKey')
                .setValue(this.plugin.settings.imageAppKey)
                .onChange(async (value) => {
                    this.plugin.settings.imageAppKey = value;
                    await this.plugin.saveSettings();
                }));
        // 图片 appSecret
        new Setting(contentEl)
            .setName('图片 appSecret')
            .setDesc('你的有道图片翻译应用密钥')
            .addText(text => text
                .setPlaceholder('imageAppSecret')
                .setValue(this.plugin.settings.imageAppSecret)
                .onChange(async (value) => {
                    this.plugin.settings.imageAppSecret = value;
                    await this.plugin.saveSettings();
                }));
        // 多行文本/图片翻译速度
        const speedDiv = contentEl.createDiv();
        speedDiv.style.margin = '16px 0';
        speedDiv.style.display = 'flex';
        speedDiv.style.alignItems = 'center';
        speedDiv.style.gap = '8px';
        speedDiv.createEl('span', { text: '多行文本/图片翻译速度（间隔ms，越小越快，API风控风险越高）:' });
        const speedInput = speedDiv.createEl('input');
        speedInput.type = 'range';
        speedInput.min = '50';
        speedInput.max = '1000';
        speedInput.step = '10';
        speedInput.value = String(this.plugin.settings.sleepInterval ?? 250);
        speedInput.style.width = '200px';
        const speedVal = speedDiv.createEl('span', { text: speedInput.value });
        speedInput.oninput = () => {
            speedVal.textContent = speedInput.value;
        };
        speedInput.onchange = async () => {
            this.plugin.settings.sleepInterval = parseInt(speedInput.value);
            await this.plugin.saveSettings();
        };
    }
}

// 在文件末尾添加 CompositeTranslateConfigModal 类
class CompositeTranslateConfigModal extends Modal {
    plugin: any;
    constructor(app: App, plugin: any) {
        super(app);
        this.plugin = plugin;
    }
    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        contentEl.createEl('h2', { text: '复合翻译设置（微软机器翻译）' });
        // 微软密钥
        new Setting(contentEl)
            .setName('微软密钥')
            .setDesc('你的微软翻译密钥')
            .addText(text => text
                .setPlaceholder('Microsoft Key')
                .setValue(this.plugin.settings.microsoftKey || '')
                .onChange(async (value) => {
                    this.plugin.settings.microsoftKey = value;
                    await this.plugin.saveSettings();
                }));
        // 微软位置/区域
        new Setting(contentEl)
            .setName('微软位置/区域')
            .setDesc('如 eastasia、global、westeurope')
            .addText(text => text
                .setPlaceholder('Microsoft Region')
                .setValue(this.plugin.settings.microsoftRegion || '')
                .onChange(async (value) => {
                    this.plugin.settings.microsoftRegion = value;
                    await this.plugin.saveSettings();
                }));
        // 微软终结点
        new Setting(contentEl)
            .setName('微软终结点')
            .setDesc('如 https://api.cognitive.microsofttranslator.com/')
            .addText(text => text
                .setPlaceholder('Microsoft Endpoint')
                .setValue(this.plugin.settings.microsoftEndpoint || '')
                .onChange(async (value) => {
                    this.plugin.settings.microsoftEndpoint = value;
                    await this.plugin.saveSettings();
                }));
        // 微软多行翻译速度
        const speedDiv = contentEl.createDiv();
        speedDiv.style.margin = '16px 0';
        speedDiv.style.display = 'flex';
        speedDiv.style.alignItems = 'center';
        speedDiv.style.gap = '8px';
        speedDiv.createEl('span', { text: '多行文本/图片翻译速度（间隔ms，越小越快，API风控风险越高）:' });
        const speedInput = speedDiv.createEl('input');
        speedInput.type = 'range';
        speedInput.min = '50';
        speedInput.max = '1000';
        speedInput.step = '10';
        speedInput.value = String(this.plugin.settings.microsoftSleepInterval ?? 250);
        speedInput.style.width = '200px';
        const speedVal = speedDiv.createEl('span', { text: speedInput.value });
        speedInput.oninput = () => {
            speedVal.textContent = speedInput.value;
        };
        speedInput.onchange = async () => {
            this.plugin.settings.microsoftSleepInterval = parseInt(speedInput.value);
            await this.plugin.saveSettings();
        };
        // 图片翻译-微软密钥
        new Setting(contentEl)
            .setName('图片翻译微软密钥')
            .setDesc('用于图片翻译的微软密钥')
            .addText(text => text
                .setPlaceholder('Microsoft Image Key')
                .setValue(this.plugin.settings.microsoftImageKey || '')
                .onChange(async (value) => {
                    this.plugin.settings.microsoftImageKey = value;
                    await this.plugin.saveSettings();
                }));
        // 图片翻译-微软位置/区域
        new Setting(contentEl)
            .setName('图片翻译微软位置/区域')
            .setDesc('用于图片翻译的微软区域，如 eastasia')
            .addText(text => text
                .setPlaceholder('Microsoft Image Region')
                .setValue(this.plugin.settings.microsoftImageRegion || '')
                .onChange(async (value) => {
                    this.plugin.settings.microsoftImageRegion = value;
                    await this.plugin.saveSettings();
                }));
        // 图片翻译-微软终结点
        new Setting(contentEl)
            .setName('图片翻译微软终结点')
            .setDesc('用于图片翻译的微软终结点')
            .addText(text => text
                .setPlaceholder('Microsoft Image Endpoint')
                .setValue(this.plugin.settings.microsoftImageEndpoint || '')
                .onChange(async (value) => {
                    this.plugin.settings.microsoftImageEndpoint = value;
                    await this.plugin.saveSettings();
                }));
    }
}

// 2. 新增 MicrosoftAdapter
class MicrosoftAdapter implements TranslateService {
    key: string;
    region: string;
    endpoint: string;
    app: App;
    constructor(key: string, region: string, endpoint: string, app: App) {
        this.key = key;
        this.region = region;
        this.endpoint = endpoint;
        this.app = app;
    }
    async translateText(text: string, from: string, to: string): Promise<string | null> {
        if (!this.key || !this.region || !this.endpoint) {
            console.warn('[MicrosoftAdapter] 缺少API参数', { key: this.key, region: this.region, endpoint: this.endpoint });
            return null;
        }
        // 修正from=auto时不加from参数，to用zh-Hans
        let url = this.endpoint.replace(/\/$/, '') + '/translate?api-version=3.0';
        if (from && from !== 'auto') {
            url += '&from=' + encodeURIComponent(from);
        }
        let toParam = to === 'zh-CHS' ? 'zh-Hans' : to;
        url += '&to=' + encodeURIComponent(toParam);
        try {
            console.log('[MicrosoftAdapter] 请求参数', { url, key: this.key.slice(0, 6) + '***', region: this.region, from, to, text });
            const resp = await fetch(url, {
                method: 'POST',
                headers: {
                    'Ocp-Apim-Subscription-Key': this.key,
                    'Ocp-Apim-Subscription-Region': this.region,
                    'Content-Type': 'application/json'
                },
                body: JSON.stringify([{ Text: text }])
            });
            const data = await resp.json();
            console.log('[MicrosoftAdapter] 响应内容', data);
            if (Array.isArray(data) && data[0] && data[0].translations && data[0].translations[0]) {
                console.log('[MicrosoftAdapter] 最终返回:', data[0].translations[0].text);
                return data[0].translations[0].text;
            }
            return null;
        } catch (e) {
            console.error('[MicrosoftAdapter] 翻译请求异常', e);
            return null;
        }
    }
    // 批量翻译
    translateTextBatch = async function (
        lines: string[],
        from: string,
        to: string
    ): Promise<string[]> {
        const endpoint = this.endpoint.endsWith('/') ? this.endpoint : this.endpoint + '/';
        // 修正from=auto时不加from参数，to用zh-Hans
        let url = `${endpoint}translate?api-version=3.0`;
        if (from && from !== 'auto') {
            url += `&from=${encodeURIComponent(from)}`;
        }
        let toParam = to === 'zh-CHS' ? 'zh-Hans' : to;
        url += `&to=${encodeURIComponent(toParam)}`;
        const body = lines.map(text => ({ Text: text }));
        const res = await fetch(url, {
            method: 'POST',
            headers: {
                'Ocp-Apim-Subscription-Key': this.key,
                'Ocp-Apim-Subscription-Region': this.region,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify(body)
        });
        if (!res.ok) throw new Error('翻译请求失败');
        const data = await res.json();
        return data.map((item: any) => item.translations[0].text);
    };
}

// 3. 微软多行文本翻译批量处理
async function unifiedBatchTranslateForMicrosoft({
    app,
    items,
    getTargetLang,
    msAdapter,
    color,
    mode,
    sleepInterval,
    plugin
}: {
    app: App,
    items: string[],
    getTargetLang: (text: string) => string,
    msAdapter: MicrosoftAdapter,
    color: string,
    mode: string,
    sleepInterval: number,
    plugin: any
}) {
    let loading = document.createElement('div');
    loading.className = 'youdao-loading';
    loading.innerHTML = `<div class='youdao-spinner'></div> 翻译正在进行，请稍等...<div class='youdao-progress' style='margin-top:16px;font-size:1.1em;'></div>`;
    Object.assign(loading.style, {
      position: 'fixed', left: '50%', top: '30%', transform: 'translate(-50%, -50%)',
      zIndex: 99999, background: '#fff', padding: '32px 48px', borderRadius: '12px',
      boxShadow: '0 2px 16px #0002', fontSize: '1.2em', textAlign: 'center'
    });
    document.body.appendChild(loading);
    const progressEl = loading.querySelector('.youdao-progress');
    try {
        let translatedArr: string[] = [];
        let successCount = 0;
        for (let i = 0; i < items.length; i++) {
            const text = (items[i] || '').trim();
            if (!text) {
                translatedArr.push('');
                if (progressEl) progressEl.textContent = `已完成${i+1}/${items.length}`;
                continue;
            }
            // 优化跳过逻辑
            const mainLang = detectMainLangSmart(text);
            let from = mainLang === 'zh' ? 'zh-Hans' : 'en'; // <-- 只改这里
            let tLang = getTargetLang(text);
            // 只在"全中文且目标语言中文"或"全英文且目标语言英文"才跳过
            const isAllZh = /^[\u4e00-\u9fa5\s，。！？、；："'（）【】《》…—·]*$/.test(text);
            const isAllEn = /^[a-zA-Z0-9\s.,!?;:'"()\[\]{}<>@#$%^&*_+=|\\/-]*$/.test(text);
            if ((tLang.startsWith('zh') && mainLang === 'zh' && isAllZh) ||
                (tLang.startsWith('en') && mainLang === 'en' && isAllEn)) {
                translatedArr.push(text);
                successCount++;
                if (progressEl) progressEl.textContent = `已完成${i+1}/${items.length}`;
                await sleep(sleepInterval ?? 250);
                console.log('[unifiedBatchTranslateForMicrosoft] 跳过翻译，原文与目标语言一致:', text);
                continue;
            }

            let translated = await msAdapter.translateText(text, from, tLang);
            if (translated && translated.trim() && translated.trim() !== text) {
                translatedArr.push(translated);
                successCount++;
            } else {
                translatedArr.push('（没有基本释义）');
            }
            if (progressEl) progressEl.textContent = `已完成${i+1}/${items.length}`;
            await sleep(sleepInterval ?? 250);
        }
        let original = items.join('\n');
        let translated = translatedArr.join('\n');
        new TranslateResultModal(app, original, translated, color, mode, plugin, successCount, items.length).open();
    } finally {
        if (loading) loading.remove();
    }
}

// 微软图片OCR识别所有行
async function microsoftOcrRecognizeLines(arrayBuffer: ArrayBuffer, msAdapter: MicrosoftAdapter): Promise<string[]> {
    const endpoint = msAdapter.endpoint.endsWith('/') ? msAdapter.endpoint : msAdapter.endpoint + '/';
    const url = `${endpoint}vision/v3.2/read/analyze`;
    const res = await fetch(url, {
        method: 'POST',
        headers: {
            'Ocp-Apim-Subscription-Key': msAdapter.key,
            'Content-Type': 'application/octet-stream'
        },
        body: arrayBuffer
    });
    if (!res.ok) throw new Error('OCR请求失败');
    const operationLocation = res.headers.get('operation-location');
    if (!operationLocation) throw new Error('未获取到OCR操作地址');
    // 轮询获取结果
    for (let i = 0; i < 20; i++) {
        await sleep(1000);
        const poll = await fetch(operationLocation, {
            headers: { 'Ocp-Apim-Subscription-Key': msAdapter.key }
        });
        const data = await poll.json();
        if (data.status === 'succeeded') {
            // 提取所有行
            const lines: string[] = [];
            for (const readResult of data.analyzeResult.readResults) {
                for (const line of readResult.lines) {
                    lines.push(line.text);
                }
            }
            return lines;
        }
        if (data.status === 'failed') throw new Error('OCR识别失败');
    }
    throw new Error('OCR超时');
}

async function microsoftTranslateImageFile(
    app: App,
    file: TFile,
    msAdapter: MicrosoftAdapter,
    color: string,
    mode: string,
    plugin: any
) {
    try {
        // 1. 读取图片为ArrayBuffer
        const arrayBuffer = await app.vault.readBinary(file);

        // 2. OCR识别所有行
        const settings = plugin.settings;
        const imageKey = settings.microsoftImageKey || '';
        const imageRegion = settings.microsoftImageRegion || '';
        const imageEndpoint = settings.microsoftImageEndpoint || '';
        if (!imageKey || !imageRegion || !imageEndpoint) {
            new Notice('请在复合翻译设置中填写"图片翻译微软密钥/区域/终结点"');
            return;
        }
        const visionAdapter = new MicrosoftAdapter(imageKey, imageRegion, imageEndpoint, app);
        let ocrLines: string[] = [];
        try {
            ocrLines = await microsoftOcrRecognizeLines(arrayBuffer, visionAdapter);
        } catch (ocrErr) {
            new Notice('OCR识别失败: ' + ocrErr.message);
            return;
        }
        if (!ocrLines.length) {
            new Notice('未识别到图片文字');
            return;
        }

        // 3. 分块翻译+进度UI，直接调用unifiedBatchTranslateForMicrosoft
        const transKey = settings.microsoftKey || '';
        const transRegion = settings.microsoftRegion || '';
        const transEndpoint = settings.microsoftEndpoint || '';
        const targetLang = settings.imageTargetLang || settings.textTargetLang || 'zh-Hans';
        const msAdapter2 = new MicrosoftAdapter(transKey, transRegion, transEndpoint, app);
        await unifiedBatchTranslateForMicrosoft({
            app,
            items: ocrLines,
            getTargetLang: (text) => detectMainLangSmart(text) === 'zh' ? 'en' : 'zh-Hans',
            msAdapter: msAdapter2,
            color,
            mode,
            sleepInterval: settings.microsoftSleepInterval ?? 250,
            plugin
        });
    } catch (e) {
        new Notice('微软图片翻译失败: ' + e.message);
    }
}

class OtherFunctionModal extends Modal {
    plugin: any;
    constructor(app: App, plugin: any) {
        super(app);
        this.plugin = plugin;
    }
    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        // 返回按钮
        const backBtn = contentEl.createEl('button', { text: '返回' });
        backBtn.style.position = 'absolute';
        backBtn.style.left = '16px';
        backBtn.style.top = '16px';
        backBtn.onclick = () => this.close();
        // 标题
        contentEl.createEl('h2', { text: '其他功能' });
        // 翻译表格按钮
        const tableBtn = contentEl.createEl('button', { text: '翻译表格' });
        tableBtn.style.display = 'block';
        tableBtn.style.margin = '40px auto';
        tableBtn.onclick = () => {
            this.close();
            new TableTranslateModal(this.app, this.plugin).open();
        };
        // 新增：表格文件和sqlite3操作按钮
        const sqliteBtn = contentEl.createEl('button', { text: '表格文件和sqlite3操作' });
        sqliteBtn.style.display = 'block';
        sqliteBtn.style.margin = '20px auto';
        sqliteBtn.onclick = () => {
            this.close();
            new TableSqliteOpModal(this.app, this.plugin).open();
        };
    }
}

class TableTranslateModal extends Modal {
    plugin: any;
    constructor(app: App, plugin: any) {
        super(app);
        this.plugin = plugin;
    }
    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        // 返回按钮
        const backBtn = contentEl.createEl('button', { text: '返回' });
        backBtn.style.position = 'absolute';
        backBtn.style.left = '16px';
        backBtn.style.top = '16px';
        backBtn.onclick = () => {
            this.close();
            new OtherFunctionModal(this.app, this.plugin).open();
        };
        // 标题
        contentEl.createEl('h2', { text: '翻译表格' });
        // 指定行按钮
        const rowBtn = contentEl.createEl('button', { text: '翻译表格指定行' });
        rowBtn.style.display = 'block';
        rowBtn.style.margin = '30px auto 10px auto';
        rowBtn.onclick = () => {
            this.close();
            new TableRowTranslateModal(this.app, this.plugin).open();
        };
        // 指定列按钮
        const colBtn = contentEl.createEl('button', { text: '翻译表格指定列' });
        colBtn.style.display = 'block';
        colBtn.style.margin = '10px auto';
        colBtn.onclick = () => {
            this.close();
            new TableColTranslateModal(this.app, this.plugin).open();
        };
    }
}

class TableRowTranslateModal extends Modal {
    private plugin: YoudaoTranslatePlugin;
    private tablePreviewEl: HTMLElement | null = null;
    private tableData: { headers: string[], rows: string[][] } | null = null;
    private translatedTableData: { headers: string[], rows: string[][] } | null = null;
    private currentPage: number = 1;
    private pageSize: number = 20;
    private totalPages: number = 1;
    private highlightRows: number[] = [];
    private rowInputEl: HTMLInputElement | null = null;
    private _originFileName: string = '';
    private showTranslated: boolean = false;
    private targetLangMode: 'auto' | 'zh' | 'en' = 'auto'; // 新增，auto=中英互译
    private modelMode: 'youdao' | 'composite'; // 新增

    constructor(app: App, plugin: YoudaoTranslatePlugin) {
        super(app);
        this.plugin = plugin;
        // 类型安全初始化
        const mode = plugin.settings.translateModel;
        this.modelMode = mode === 'composite' ? 'composite' : 'youdao';
    }

    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        // 返回按钮，样式与 OtherFunctionModal 一致
        const backBtn = contentEl.createEl('button', { text: '返回' });
        backBtn.style.position = 'absolute';
        backBtn.style.left = '16px';
        backBtn.style.top = '16px';
        backBtn.onclick = () => {
            this.close();
            new TableTranslateModal(this.app, this.plugin).open();
        };
        contentEl.createEl('h2', { text: '翻译表格指定行' });
        contentEl.createEl('p', { text: '请选择要翻译的表格文件' });

        // --- 控件区：按钮和输入框同一行 ---
        const controlBar = contentEl.createEl('div');
        controlBar.style.display = 'flex';
        controlBar.style.alignItems = 'center';
        controlBar.style.gap = '12px';
        controlBar.style.marginBottom = '10px';

        // 1. 插入表格按钮
        const insertBtn = controlBar.createEl('button', { text: '插入表格' });
        insertBtn.onclick = () => this.showFilePicker();

        // 2. 翻译模型按钮
        const modelBtn = controlBar.createEl('button', { text: '翻译模型' });
        modelBtn.textContent = this.modelMode === 'composite' ? '复合翻译' : '有道翻译';
        modelBtn.onclick = () => showDropdown(modelBtn, ['有道翻译', '复合翻译'], val => {
            modelBtn.textContent = val;
            this.modelMode = val === '复合翻译' ? 'composite' : 'youdao';
        });

        // 3. 翻译语言按钮
        const langBtn = controlBar.createEl('button', { text: '翻译语言' });
        langBtn.textContent = this.targetLangMode === 'auto' ? '中英互译' : (this.targetLangMode === 'zh' ? '目标语言为中文' : '目标语言为英文');
        langBtn.onclick = () => showDropdown(langBtn, ['中英互译', '目标语言为中文', '目标语言为英文'], val => {
            langBtn.textContent = val;
            if (val === '中英互译') this.targetLangMode = 'auto';
            else if (val === '目标语言为中文') this.targetLangMode = 'zh';
            else this.targetLangMode = 'en';
        });

        // 4. 选定行翻译输入框
        const rowInput = controlBar.createEl('input', { type: 'text', placeholder: '选定行翻译' });
        rowInput.style.flex = '1';
        rowInput.style.minWidth = '120px';
        rowInput.style.marginLeft = '8px';
        this.rowInputEl = rowInput;
        rowInput.value = this.highlightRows.join('\\');
        rowInput.oninput = () => {
            const input = rowInput.value.replace(/\s/g, '');
            this.highlightRows = input.split('\\').map(Number).filter(n => !isNaN(n));
            this.renderTablePreview();
        };

        // 5. 表格预览区
        this.tablePreviewEl = contentEl.createEl('div');
    }

    onClose() {
        const { contentEl } = this;
        contentEl.empty();
    }
    
    private showFilePicker() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.csv,.tsv,.md,.xlsx,.ods,.sqlite3';
        input.onchange = async () => {
            const file = input.files?.[0];
            if (!file) return;
            const ext = file.name.split('.').pop()?.toLowerCase();
            if (file.name.endsWith('.sqlite3')) {
                await this.parseSqlite(file);
            } else if (ext === 'xlsx' || ext === 'ods') {
                const arrayBuffer = await file.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                // 获取实际表格范围，生成完整二维数组
                const ref = worksheet['!ref'];
                if (!ref) {
                    new Notice('该表为空或无数据');
                    return;
                }
                const range = XLSX.utils.decode_range(ref);
                const maxRow = range.e.r;
                const maxCol = range.e.c;
                const rows: string[][] = [];
                for (let r = 0; r <= maxRow; r++) {
                    const row: string[] = [];
                    for (let c = 0; c <= maxCol; c++) {
                        const cellAddress = XLSX.utils.encode_cell({ c, r });
                        const cell = worksheet[cellAddress];
                        row.push(cell ? String(cell.v) : '');
                    }
                    rows.push(row);
                }
                const headers = rows.length > 0 ? rows[0] : [];
                const dataRows = rows.slice(1);
                this.tableData = { headers, rows: dataRows };
                this.currentPage = 1;
                this.pageSize = 20;
                this.totalPages = this.tableData.rows.length > 0 ? Math.ceil(this.tableData.rows.length / this.pageSize) : 1;
                this.renderTablePreview();
            } else {
                const text = await file.text();
                if (ext === 'tsv') {
                    this.parseTable(text, false, '\t');
                } else {
                    this.parseTable(text, ext === 'md');
                }
            }
        };
        input.click();
    }
    private async exportTableFile(format: string) {
        const data = this.showTranslated && this.translatedTableData ? this.translatedTableData : this.tableData;
        if (!data) {
            new Notice('无表格数据，无法导出');
            return;
        }
        // 1. 获取原始文件名或让用户输入
        let defaultName = this._originFileName || '表格名';
        new FileNamePromptModal(this.app, defaultName + '_translate', (fileName) => {
            if (!fileName) return;
            let ext = format;
            if (fileName.endsWith('.' + ext)) fileName = fileName.slice(0, -1 - ext.length);
            let finalName = fileName + '.' + ext;
            const headers = data.headers;
            const rows = data.rows.map(row => [...row]);
            let blob: Blob;
            if (format === 'csv' || format === 'tsv') {
                const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                const csv = XLSX.utils.sheet_to_csv(ws, { FS: format === 'tsv' ? '\t' : ',' });
                blob = new Blob([csv], { type: 'text/' + format });
            } else if (format === 'md') {
                const md = markdownTable([headers, ...rows]);
                blob = new Blob([md], { type: 'text/markdown' });
            } else if (format === 'xlsx' || format === 'ods') {
                const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                const wbout = XLSX.write(wb, { bookType: format, type: 'array' });
                blob = new Blob([wbout], { type: 'application/octet-stream' });
            } else {
                new Notice('暂不支持该格式导出');
                return;
            }
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = finalName;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                URL.revokeObjectURL(a.href);
                a.remove();
            }, 100);
            new Notice('导出成功: ' + finalName);
        }).open();
    }

    private async parseSqlite(file: File) {
        const arrayBuffer = await file.arrayBuffer();
        const initSqlJs = await loadSqlJs();
        const SQL = await initSqlJs({ locateFile: (file: string) => 'https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/' + file });
        const db = new SQL.Database(new Uint8Array(arrayBuffer));
        const tableNames = db.exec('SELECT name FROM sqlite_master WHERE type="table"')[0].values.flat();
        showDropdown(this.tablePreviewEl!, tableNames, async tableName => {
            const tableData = db.exec(`SELECT * FROM ${tableName}`)[0];
            this.tableData = { headers: tableData.columns, rows: tableData.values };
            this.currentPage = 1;
            this.pageSize = 20;
            this.totalPages = this.tableData.rows.length > 0 ? Math.ceil(this.tableData.rows.length / this.pageSize) : 1;
            this.renderTablePreview();
        });
    }

    private parseTable(text: string, isMarkdown: boolean, delimiter: string = ',') {
        this.tableData = isMarkdown ? parseMarkdownTable(text) : parseCsvTsv(text, delimiter);
        this.currentPage = 1;
        this.pageSize = 20;
        this.totalPages = this.tableData.rows.length > 0 ? Math.ceil(this.tableData.rows.length / this.pageSize) : 1;
        this.renderTablePreview();
    }

    private renderTablePreview() {
        if (!this.tablePreviewEl || !this.tableData) return;
        this.tablePreviewEl.empty();
        // 添加右上角切换按钮，与TableColTranslateModal一致
        const switchBtn = document.createElement('button');
        switchBtn.textContent = this.showTranslated ? '查看原始表格' : '查看翻译后表格';
        switchBtn.style.float = 'right';
        switchBtn.onclick = () => {
            this.showTranslated = !this.showTranslated;
            this.renderTablePreview();
        };
        this.tablePreviewEl.appendChild(switchBtn);
        // 渲染表格
        const data = this.showTranslated && this.translatedTableData ? this.translatedTableData : this.tableData;
        if (!data) return;
        // 分页渲染
        const start = (this.currentPage - 1) * this.pageSize;
        const end = start + this.pageSize;
        const pageRows = data.rows.slice(start, end);
        this.tablePreviewEl.appendChild(renderHtmlTable(
            data.headers,
            pageRows,
            this.highlightRows,
            [],
            (row) => {
                if (this.highlightRows.includes(row)) {
                    this.highlightRows = this.highlightRows.filter(r => r !== row);
                } else {
                    this.highlightRows.push(row);
                }
                // 同步输入框内容
                if (this.rowInputEl) this.rowInputEl.value = this.highlightRows.join('\\');
                this.renderTablePreview();
            },
            () => {},
            true, // enableRowClick
            start // startIndex
        ));
        this.renderPagination();
    }

    private renderPagination() {
        if (!this.tablePreviewEl || !this.tableData) return;
        let oldBar = this.tablePreviewEl.querySelector('.pagination-bar') as HTMLElement;
        if (oldBar) oldBar.remove();
        const bar = document.createElement('div');
        bar.className = 'pagination-bar';
        bar.style.display = 'flex';
        bar.style.justifyContent = 'space-between';
        bar.style.alignItems = 'center';
        bar.style.margin = '12px 0 0 0';
        bar.style.width = '100%';

        // 分页控件
        const pagDiv = document.createElement('div');
        pagDiv.className = 'pagination';
        pagDiv.style.display = 'flex';
        pagDiv.style.justifyContent = 'flex-start';
        pagDiv.style.alignItems = 'center';
        pagDiv.style.gap = '8px';
        pagDiv.style.margin = '0';

        // 页码信息
        const infoSpan = document.createElement('span');
        infoSpan.textContent = `第 ${this.currentPage} / ${this.totalPages} 页`;
        infoSpan.style.marginRight = '12px';
        pagDiv.appendChild(infoSpan);

        // 页码输入框
        const pageInput = document.createElement('input');
        pageInput.type = 'number';
        pageInput.min = '1';
        pageInput.max = String(this.totalPages);
        pageInput.value = String(this.currentPage);
        pageInput.style.width = '48px';
        pageInput.style.marginRight = '12px';
        pageInput.style.fontSize = '1em';
        pageInput.style.verticalAlign = 'middle';
        pageInput.title = '跳转到指定页';
        pageInput.onkeydown = (e) => {
            if (e.key === 'Enter') {
                let val = parseInt(pageInput.value);
                if (isNaN(val) || val < 1) val = 1;
                if (val > this.totalPages) val = this.totalPages;
                if (val !== this.currentPage) {
                    this.currentPage = val;
                    this.renderTablePreview();
                }
            }
        };
        pageInput.onblur = () => {
            let val = parseInt(pageInput.value);
            if (isNaN(val) || val < 1) val = 1;
            if (val > this.totalPages) val = this.totalPages;
            if (val !== this.currentPage) {
                this.currentPage = val;
                this.renderTablePreview();
            }
        };
        pagDiv.appendChild(pageInput);

        // 上一页
        const prevBtn = document.createElement('button');
        prevBtn.textContent = '上一页';
        prevBtn.disabled = this.currentPage === 1;
        prevBtn.onclick = () => {
            if (this.currentPage > 1) {
                this.currentPage--;
                this.renderTablePreview();
            }
        };
        pagDiv.appendChild(prevBtn);

        // 下一页
        const nextBtn = document.createElement('button');
        nextBtn.textContent = '下一页';
        nextBtn.disabled = this.currentPage === this.totalPages;
        nextBtn.onclick = () => {
            if (this.currentPage < this.totalPages) {
                this.currentPage++;
                this.renderTablePreview();
            }
        };
        pagDiv.appendChild(nextBtn);

        // 翻译按钮
        const translateBtn = document.createElement('button');
        translateBtn.textContent = '翻译 ▼';
        translateBtn.style.background = '#1a73e8';
        translateBtn.style.color = '#fff';
        translateBtn.style.fontSize = '1.1em';
        translateBtn.style.padding = '10px 32px';
        translateBtn.style.border = 'none';
        translateBtn.style.borderRadius = '8px';
        translateBtn.style.boxShadow = '0 2px 8px #0002';
        translateBtn.style.cursor = 'pointer';
        translateBtn.style.position = 'relative';
        translateBtn.onmouseenter = () => translateBtn.style.background = '#1765c1';
        translateBtn.onmouseleave = () => translateBtn.style.background = '#1a73e8';
        let dropdown: HTMLDivElement | null = null;
        translateBtn.onclick = () => {
            if (dropdown) { dropdown.remove(); dropdown = null; return; }
            dropdown = document.createElement('div');
            dropdown.style.position = 'absolute';
            dropdown.style.top = '100%';
            dropdown.style.right = '0';
            dropdown.style.background = '#fff';
            dropdown.style.border = '1px solid #ccc';
            dropdown.style.borderRadius = '8px';
            dropdown.style.boxShadow = '0 2px 8px #0002';
            dropdown.style.minWidth = '220px';
            dropdown.style.padding = '8px 0';
            dropdown.style.zIndex = '10001';
            dropdown.style.maxHeight = '60vh';
            dropdown.style.overflowY = 'auto';
            const options = [
                { label: '翻译结果覆盖原文件', action: async () => { dropdown?.remove(); dropdown = null; new Notice('已覆盖原文件（模拟，实际功能待实现）'); }},
                { label: '生成副本文件', action: async () => { dropdown?.remove(); dropdown = null; new Notice('副本文件已保存（模拟，实际功能待实现）'); }},
                { label: '导出为自定义类型文件', action: async () => {
                    dropdown?.remove(); dropdown = null;
                    const types = ['csv','tsv','md','xlsx','ods'];
                    showDropdown(translateBtn, types.map(t=>'.'+t), async (ext) => {
                        let defaultName = this._originFileName || '表格名';
                        const modal = new FileNamePromptModal(this.app, defaultName + '_translate', async (fileName) => {
                            if (!fileName) return;
                            // 1. 显示 loading
                            let loading = document.createElement('div');
                            loading.className = 'youdao-loading';
                            loading.innerHTML = `<div class='youdao-spinner'></div> 正在翻译并导出，请稍候...<div class='progress-info'></div>`;
                            Object.assign(loading.style, {
                                position: 'fixed', left: '50%', top: '30%', transform: 'translate(-50%, -50%)',
                                zIndex: 99999, background: '#fff', padding: '32px 48px', borderRadius: '12px',
                                boxShadow: '0 2px 16px #0002', fontSize: '1.2em', textAlign: 'center'
                            });
                            document.body.appendChild(loading);
                            const progressEl = loading.querySelector('.progress-info');
                            try {
                                // 深拷贝原始数据
                                const headers = [...this.tableData!.headers];
                                const rows = this.tableData!.rows.map(row => [...row]);
                                let totalCells = 0, translatedCells = 0;
                                for (const rowNum of this.highlightRows) {
                                    const rowIdx = rowNum - 1;
                                    if (rowIdx < 0 || rowIdx >= rows.length) continue;
                                    totalCells += rows[rowIdx].length;
                                }
                                for (const rowNum of this.highlightRows) {
                                    const rowIdx = rowNum - 1;
                                    if (rowIdx < 0 || rowIdx >= rows.length) continue;
                                    for (let colIdx = 0; colIdx < rows[rowIdx].length; colIdx++) {
                                        const translatedCell = await this.translateCell(rows[rowIdx][colIdx]);
                                        if (translatedCell === rows[rowIdx][colIdx]) {
                                            translatedCells++;
                                            if (progressEl) progressEl.textContent = `已翻译${translatedCells}/${totalCells}单元格`;
                                        } else {
                                            rows[rowIdx][colIdx] = translatedCell;
                                            translatedCells++;
                                            if (progressEl) progressEl.textContent = `已翻译${translatedCells}/${totalCells}单元格`;
                                            await sleep(this.plugin.settings.translateModel === 'composite'
                                                ? this.plugin.settings.microsoftSleepInterval ?? 250
                                                : this.plugin.settings.sleepInterval ?? 250);
                                        }
                                    }
                                }
                                // 只赋值副本
                                this.translatedTableData = { headers, rows };
                                this.showTranslated = true;
                                this.renderTablePreview();
                
                                // 导出
                                let ext2 = ext.replace('.', '');
                                let finalName = fileName.endsWith('.' + ext2) ? fileName : (fileName + '.' + ext2);
                                let blob: Blob;
                                if (ext2 === 'csv' || ext2 === 'tsv') {
                                    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                                    const csv = XLSX.utils.sheet_to_csv(ws, { FS: ext2 === 'tsv' ? '\t' : ',' });
                                    blob = new Blob([csv], { type: 'text/' + ext2 });
                                } else if (ext2 === 'md') {
                                    const md = markdownTable([headers, ...rows]);
                                    blob = new Blob([md], { type: 'text/markdown' });
                                } else if (ext2 === 'xlsx' || ext2 === 'ods') {
                                    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                                    const wb = XLSX.utils.book_new();
                                    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                                    const wbout = XLSX.write(wb, { bookType: ext2, type: 'array' });
                                    blob = new Blob([wbout], { type: 'application/octet-stream' });
                                } else {
                                    new Notice('暂不支持该格式导出');
                                    return;
                                }
                                const a = document.createElement('a');
                                a.href = URL.createObjectURL(blob);
                                a.download = finalName;
                                document.body.appendChild(a);
                                a.click();
                                setTimeout(() => {
                                    URL.revokeObjectURL(a.href);
                                    a.remove();
                                }, 100);
                                new Notice('导出成功: ' + finalName);
                            } catch (e) {
                                new Notice('导出失败: ' + (e?.message || e));
                            } finally {
                                if (loading) loading.remove();
                            }
                        });
                        modal.open();
                    });
                }},
            ];
            options.forEach(opt => {
                const item = document.createElement('div');
                item.textContent = opt.label;
                item.style.padding = '12px 24px';
                item.style.cursor = 'pointer';
                item.style.background = '#fff'; // 默认白底
                item.style.color = '#222';      // 默认黑字
                item.onmouseenter = () => {
                    item.style.background = '#e3eafc'; // 悬停浅蓝底
                    item.style.color = '#1765c1';      // 悬停蓝字
                };
                item.onmouseleave = () => {
                    item.style.background = '#fff';
                    item.style.color = '#222';
                };
                item.onclick = (event) => {
                    event.stopPropagation();
                    if (dropdown) dropdown.remove();
                    if (typeof opt.action === 'function') opt.action();
                };
                dropdown!.appendChild(item);
            });
            translateBtn.appendChild(dropdown);
            setTimeout(() => {
                const close = (ev: MouseEvent) => {
                    if (!dropdown) return;
                    if (!dropdown.contains(ev.target as Node) && ev.target !== translateBtn) {
                        dropdown.remove();
                        dropdown = null;
                        document.removeEventListener('mousedown', close);
                    }
                };
                document.addEventListener('mousedown', close);
            }, 10);
        };

        bar.appendChild(pagDiv);
        bar.appendChild(translateBtn);
        this.tablePreviewEl.appendChild(bar);
    }

    private async translateCell(cell: string): Promise<string> {
        const settings = this.plugin.settings;
        if (!cell || !(cell ?? '').toString().trim()) return cell;
        // 判断目标语言
        let targetLang = settings.textTargetLang;
        if (this.targetLangMode === 'auto') {
            if (settings.isBiDirection) {
                targetLang = /[\u4e00-\u9fa5]/.test(cell) ? 'en' : (targetLang === 'zh' ? 'zh-CHS' : targetLang);
            }
        } else if (this.targetLangMode === 'zh') {
            targetLang = this.modelMode === 'composite' ? 'zh-Hans' : 'zh-CHS';
        } else if (this.targetLangMode === 'en') {
            targetLang = 'en';
        }
        // 新增：如果目标语言为中文，且内容包含繁体，则用 t2s 转为简体
        if ((targetLang === 'zh-CHS' || targetLang === 'zh-Hans')) {
            const simplified = t2s(cell);
            if (simplified !== cell) {
                return simplified;
            }
            // 原有逻辑：如果内容全是简体中文，直接跳过
            if (/^[\u4e00-\u9fa5\s，。！？、；："'（）【】《》…—·]*$/.test(cell)) {
                return cell;
            }
        }
        if (targetLang === 'en' && /^[a-zA-Z0-9\s.,!?;:'"()\[\]{}<>@#$%^&*_+=|\\/-]*$/.test(cell)) {
            return cell;
        }
        // 选择翻译模型
        if (this.modelMode === 'composite') {
            // 微软翻译
            const msAdapter = new MicrosoftAdapter(
                settings.microsoftKey || '',
                settings.microsoftRegion || '',
                settings.microsoftEndpoint || '',
                this.app
            );
            const from = /[\u4e00-\u9fa5]/.test(cell) ? 'zh-Hans' : 'en';
            const tLang = targetLang === 'zh' ? 'zh-Hans' : targetLang;
            const translated = await msAdapter.translateText(cell, from, tLang);
            console.log('[翻译日志][MicrosoftAdapter] 原文:', cell, 'from:', from, 'to:', tLang, '返回:', translated);
            return translated || cell;
        } else {
            // 有道翻译
            const translator: TranslateService = new YoudaoAdapter(
                settings.textAppKey,
                settings.textAppSecret,
                this.app,
                settings.serverPort
            );
            const translated = await translator.translateText(cell, 'auto', targetLang);
            console.log('[翻译日志][YoudaoAdapter] 原文:', cell, 'from:auto', 'to:', targetLang, '返回:', translated);
            return translated || cell;
        }
    }
}

class TableColTranslateModal extends Modal {
    private plugin: YoudaoTranslatePlugin;
    private tablePreviewEl: HTMLElement | null = null;
    private tableData: { headers: string[], rows: string[][] } | null = null;
    private translatedTableData: { headers: string[], rows: string[][] } | null = null;
    private currentPage: number = 1;
    private pageSize: number = 20;
    private totalPages: number = 1;
    private highlightCols: number[] = [];
    private colInputEl: HTMLInputElement | null = null;
    private _originFileName: string = '';
    private showTranslated: boolean = false;
    private targetLangMode: 'auto' | 'zh' | 'en' = 'auto'; // 新增，auto=中英互译
    private modelMode: 'youdao' | 'composite'; // 类型安全

    constructor(app: App, plugin: YoudaoTranslatePlugin) {
        super(app);
        this.plugin = plugin;
        // 类型安全初始化
        const mode = plugin.settings.translateModel;
        this.modelMode = mode === 'composite' ? 'composite' : 'youdao';
    }

    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        // 返回按钮
        const backBtn = contentEl.createEl('button', { text: '返回' });
        backBtn.style.position = 'absolute';
        backBtn.style.left = '16px';
        backBtn.style.top = '16px';
        backBtn.onclick = () => {
            this.close();
            new TableTranslateModal(this.app, this.plugin).open();
        };
        contentEl.createEl('h2', { text: '翻译表格指定列' });
        contentEl.createEl('p', { text: '请选择要翻译的表格文件' });

        // 控件区
        const controlBar = contentEl.createEl('div');
        controlBar.style.display = 'flex';
        controlBar.style.alignItems = 'center';
        controlBar.style.gap = '12px';
        controlBar.style.marginBottom = '10px';

        // 插入表格按钮
        const insertBtn = controlBar.createEl('button', { text: '插入表格' });
        insertBtn.onclick = () => this.showFilePicker();

        // 翻译模型按钮
        const modelBtn = controlBar.createEl('button', { text: '翻译模型' });
        modelBtn.textContent = this.modelMode === 'composite' ? '复合翻译' : '有道翻译';
        modelBtn.onclick = () => showDropdown(modelBtn, ['有道翻译', '复合翻译'], val => {
            modelBtn.textContent = val;
            this.modelMode = val === '复合翻译' ? 'composite' : 'youdao';
        });

        // 翻译语言按钮
        const langBtn = controlBar.createEl('button', { text: '翻译语言' });
        langBtn.textContent = this.targetLangMode === 'auto' ? '中英互译' : (this.targetLangMode === 'zh' ? '目标语言为中文' : '目标语言为英文');
        langBtn.onclick = () => showDropdown(langBtn, ['中英互译', '目标语言为中文', '目标语言为英文'], val => {
            langBtn.textContent = val;
            if (val === '中英互译') this.targetLangMode = 'auto';
            else if (val === '目标语言为中文') this.targetLangMode = 'zh';
            else this.targetLangMode = 'en';
        });

        // 选定列翻译输入框
        const colInput = controlBar.createEl('input', { type: 'text', placeholder: '选定列翻译' });
        colInput.style.flex = '1';
        colInput.style.minWidth = '120px';
        colInput.style.marginLeft = '8px';
        this.colInputEl = colInput;
        colInput.value = this.highlightCols.join('\\');
        colInput.oninput = () => {
            const input = colInput.value.replace(/\s/g, '');
            this.highlightCols = input.split('\\').map(Number).filter(n => !isNaN(n));
            this.renderTablePreview();
        };

        // 表格预览区
        this.tablePreviewEl = contentEl.createEl('div');
    }

    onClose() {
        this.contentEl.empty();
    }

    private showFilePicker() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.csv,.tsv,.md,.xlsx,.ods,.sqlite3';
        input.onchange = async () => {
            const file = input.files?.[0];
            if (!file) return;
            const ext = file.name.split('.').pop()?.toLowerCase();
            if (file.name.endsWith('.sqlite3')) {
                await this.parseSqlite(file);
            } else if (ext === 'xlsx' || ext === 'ods') {
                const arrayBuffer = await file.arrayBuffer();
                const workbook = XLSX.read(arrayBuffer, { type: 'array' });
                const sheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[sheetName];
                // 获取实际表格范围，生成完整二维数组
                const ref = worksheet['!ref'];
                if (!ref) {
                    new Notice('该表为空或无数据');
            return;
        }
                const range = XLSX.utils.decode_range(ref);
                const maxRow = range.e.r;
                const maxCol = range.e.c;
                const rows: string[][] = [];
                for (let r = 0; r <= maxRow; r++) {
                    const row: string[] = [];
                    for (let c = 0; c <= maxCol; c++) {
                        const cellAddress = XLSX.utils.encode_cell({ c, r });
                        const cell = worksheet[cellAddress];
                        row.push(cell ? String(cell.v) : '');
                    }
                    rows.push(row);
                }
                const headers = rows.length > 0 ? rows[0] : [];
                const dataRows = rows.slice(1);
                this.tableData = { headers, rows: dataRows };
                this.currentPage = 1;
                this.pageSize = 20;
                this.totalPages = this.tableData.rows.length > 0 ? Math.ceil(this.tableData.rows.length / this.pageSize) : 1;
                this.renderTablePreview();
            } else {
                const text = await file.text();
                if (ext === 'tsv') {
                    this.parseTable(text, false, '\t');
                } else {
                    this.parseTable(text, ext === 'md');
                }
            }
        };
        input.click();
    }

    private async parseSqlite(file: File) {
        const arrayBuffer = await file.arrayBuffer();
        const initSqlJs = await loadSqlJs();
        const SQL = await initSqlJs({ locateFile: (file: string) => 'https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/' + file });
        const db = new SQL.Database(new Uint8Array(arrayBuffer));
        const tableNames = db.exec('SELECT name FROM sqlite_master WHERE type=\"table\"')[0].values.flat();
        showDropdown(this.tablePreviewEl!, tableNames, async tableName => {
            const tableData = db.exec(`SELECT * FROM ${tableName}`)[0];
            this.tableData = { headers: tableData.columns, rows: tableData.values };
            this.currentPage = 1;
            this.pageSize = 20;
            this.totalPages = this.tableData.rows.length > 0 ? Math.ceil(this.tableData.rows.length / this.pageSize) : 1;
            this.renderTablePreview();
        });
    }

    private parseTable(text: string, isMarkdown: boolean, delimiter: string = ',') {
        this.tableData = isMarkdown ? parseMarkdownTable(text) : parseCsvTsv(text, delimiter);
        this.currentPage = 1;
        this.pageSize = 20;
        this.totalPages = this.tableData.rows.length > 0 ? Math.ceil(this.tableData.rows.length / this.pageSize) : 1;
        this.renderTablePreview();
    }


    private renderTablePreview() {
        if (!this.tablePreviewEl || !this.tableData) return;
        this.tablePreviewEl.empty();
        // 右上角切换按钮
        const switchBtn = document.createElement('button');
        switchBtn.textContent = this.showTranslated ? '查看原始表格' : '查看翻译后表格';
        switchBtn.style.float = 'right';
        switchBtn.onclick = () => {
            this.showTranslated = !this.showTranslated;
            this.renderTablePreview();
        };
        this.tablePreviewEl.appendChild(switchBtn);
        // 渲染表格
        const data = this.showTranslated && this.translatedTableData ? this.translatedTableData : this.tableData;
        if (!data) return;
        // 分页渲染
        const start = (this.currentPage - 1) * this.pageSize;
        const end = start + this.pageSize;
        const pageRows = data.rows.slice(start, end);
        this.tablePreviewEl.appendChild(renderHtmlTable(
            data.headers,
            pageRows,
            [],
            this.highlightCols,
            () => {},
            (col) => {
                if (this.highlightCols.includes(col)) {
                    this.highlightCols = this.highlightCols.filter(c => c !== col);
                } else {
                    this.highlightCols.push(col);
                }
                if (this.colInputEl) this.colInputEl.value = this.highlightCols.join('\\');
                this.renderTablePreview();
            },
            false, // enableRowClick
            start // startIndex
        ));
                this.renderPagination();
    }

    private renderPagination() {
        if (!this.tablePreviewEl || !this.tableData) return;
        let oldBar = this.tablePreviewEl.querySelector('.pagination-bar') as HTMLElement;
        if (oldBar) oldBar.remove();
        const bar = document.createElement('div');
        bar.className = 'pagination-bar';
        bar.style.display = 'flex';
        bar.style.justifyContent = 'space-between';
        bar.style.alignItems = 'center';
        bar.style.margin = '12px 0 0 0';
        bar.style.width = '100%';

        // 分页控件
        const pagDiv = document.createElement('div');
        pagDiv.className = 'pagination';
        pagDiv.style.display = 'flex';
        pagDiv.style.justifyContent = 'flex-start';
        pagDiv.style.alignItems = 'center';
        pagDiv.style.gap = '8px';
        pagDiv.style.margin = '0';

        // 页码信息
        const infoSpan = document.createElement('span');
        infoSpan.textContent = `第 ${this.currentPage} / ${this.totalPages} 页`;
        infoSpan.style.marginRight = '12px';
        pagDiv.appendChild(infoSpan);

        // 页码输入框
        const pageInput = document.createElement('input');
        pageInput.type = 'number';
        pageInput.min = '1';
        pageInput.max = String(this.totalPages);
        pageInput.value = String(this.currentPage);
        pageInput.style.width = '48px';
        pageInput.style.marginRight = '12px';
        pageInput.style.fontSize = '1em';
        pageInput.style.verticalAlign = 'middle';
        pageInput.title = '跳转到指定页';
        pageInput.onkeydown = (e) => {
            if (e.key === 'Enter') {
                let val = parseInt(pageInput.value);
                if (isNaN(val) || val < 1) val = 1;
                if (val > this.totalPages) val = this.totalPages;
                if (val !== this.currentPage) {
                    this.currentPage = val;
                    this.renderTablePreview();
                }
            }
        };
        pageInput.onblur = () => {
            let val = parseInt(pageInput.value);
            if (isNaN(val) || val < 1) val = 1;
            if (val > this.totalPages) val = this.totalPages;
            if (val !== this.currentPage) {
                this.currentPage = val;
                this.renderTablePreview();
            }
        };
        pagDiv.appendChild(pageInput);

        // 上一页
        const prevBtn = document.createElement('button');
        prevBtn.textContent = '上一页';
        prevBtn.disabled = this.currentPage === 1;
        prevBtn.onclick = () => {
            if (this.currentPage > 1) {
                this.currentPage--;
                this.renderTablePreview();
            }
        };
        pagDiv.appendChild(prevBtn);

        // 下一页
        const nextBtn = document.createElement('button');
        nextBtn.textContent = '下一页';
        nextBtn.disabled = this.currentPage === this.totalPages;
        nextBtn.onclick = () => {
            if (this.currentPage < this.totalPages) {
                this.currentPage++;
                this.renderTablePreview();
            }
        };
        pagDiv.appendChild(nextBtn);

        // 翻译按钮
        const translateBtn = document.createElement('button');
        translateBtn.textContent = '翻译 ▼';
        translateBtn.style.background = '#1a73e8';
        translateBtn.style.color = '#fff';
        translateBtn.style.fontSize = '1.1em';
        translateBtn.style.padding = '10px 32px';
        translateBtn.style.border = 'none';
        translateBtn.style.borderRadius = '8px';
        translateBtn.style.boxShadow = '0 2px 8px #0002';
        translateBtn.style.cursor = 'pointer';
        translateBtn.style.position = 'relative';
        translateBtn.onmouseenter = () => translateBtn.style.background = '#1765c1';
        translateBtn.onmouseleave = () => translateBtn.style.background = '#1a73e8';
        let dropdown: HTMLDivElement | null = null;
        translateBtn.onclick = () => {
            if (dropdown) { dropdown.remove(); dropdown = null; return; }
            dropdown = document.createElement('div');
            dropdown.style.position = 'absolute';
            dropdown.style.top = '100%';
            dropdown.style.right = '0';
            dropdown.style.background = '#fff';
            dropdown.style.border = '1px solid #ccc';
            dropdown.style.borderRadius = '8px';
            dropdown.style.boxShadow = '0 2px 8px #0002';
            dropdown.style.minWidth = '220px';
            dropdown.style.padding = '8px 0';
            dropdown.style.zIndex = '10001';
            dropdown.style.maxHeight = '60vh';
            dropdown.style.overflowY = 'auto';
            const options = [
                { label: '翻译结果覆盖原文件', action: async () => { dropdown?.remove(); dropdown = null; new Notice('已覆盖原文件（模拟，实际功能待实现）'); }},
                { label: '生成副本文件', action: async () => { dropdown?.remove(); dropdown = null; new Notice('副本文件已保存（模拟，实际功能待实现）'); }},
                { label: '导出为自定义类型文件', action: async () => {
                    dropdown?.remove(); dropdown = null;
                    const types = ['csv','tsv','md','xlsx','ods'];
                    showDropdown(translateBtn, types.map(t=>'.'+t), async (ext) => {
                        let defaultName = this._originFileName || '表格名';
                        const modal = new FileNamePromptModal(this.app, defaultName + '_translate', async (fileName) => {
                            if (!fileName) return;
                            // 1. 显示 loading
                            let loading = document.createElement('div');
                            loading.className = 'youdao-loading';
                            loading.innerHTML = `<div class='youdao-spinner'></div> 正在翻译并导出，请稍候...<div class='progress-info'></div>`;
                            Object.assign(loading.style, {
                                position: 'fixed', left: '50%', top: '30%', transform: 'translate(-50%, -50%)',
                                zIndex: 99999, background: '#fff', padding: '32px 48px', borderRadius: '12px',
                                boxShadow: '0 2px 16px #0002', fontSize: '1.2em', textAlign: 'center'
                            });
                            document.body.appendChild(loading);
                            const progressEl = loading.querySelector('.progress-info');
                            try {
                                // 深拷贝原始数据
                                const headers = [...this.tableData!.headers];
                                const rows = this.tableData!.rows.map(row => [...row]);
                                let totalCells = 0, translatedCells = 0;
                                for (const colNum of this.highlightCols) {
                                    const colIdx = colNum - 1;
                                    if (colIdx < 0 || colIdx >= headers.length) continue;
                                    totalCells += rows.length;
                                }
                                for (const colNum of this.highlightCols) {
                                    const colIdx = colNum - 1;
                                    if (colIdx < 0 || colIdx >= headers.length) continue;
                                    for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
                                        const translatedCell = await this.translateCell(rows[rowIdx][colIdx]);
                                        if (translatedCell === rows[rowIdx][colIdx]) {
                                            translatedCells++;
                                            if (progressEl) progressEl.textContent = `已翻译${translatedCells}/${totalCells}单元格`;
                                        } else {
                                            rows[rowIdx][colIdx] = translatedCell;
                                            translatedCells++;
                                            if (progressEl) progressEl.textContent = `已翻译${translatedCells}/${totalCells}单元格`;
                                            await sleep(this.plugin.settings.translateModel === 'composite'
                                                ? this.plugin.settings.microsoftSleepInterval ?? 250
                                                : this.plugin.settings.sleepInterval ?? 250);
                                        }
                                    }
                                }
                                // 只赋值副本
                                this.translatedTableData = { headers, rows };
                                this.showTranslated = true;
                                this.renderTablePreview();
                
                                // 导出
                                let ext2 = ext.replace('.', '');
                                let finalName = fileName.endsWith('.' + ext2) ? fileName : (fileName + '.' + ext2);
                                let blob: Blob;
                                if (ext2 === 'csv' || ext2 === 'tsv') {
                                    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                                    const csv = XLSX.utils.sheet_to_csv(ws, { FS: ext2 === 'tsv' ? '\t' : ',' });
                                    blob = new Blob([csv], { type: 'text/' + ext2 });
                                } else if (ext2 === 'md') {
                                    const md = markdownTable([headers, ...rows]);
                                    blob = new Blob([md], { type: 'text/markdown' });
                                } else if (ext2 === 'xlsx' || ext2 === 'ods') {
                                    const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                                    const wb = XLSX.utils.book_new();
                                    XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                                    const wbout = XLSX.write(wb, { bookType: ext2, type: 'array' });
                                    blob = new Blob([wbout], { type: 'application/octet-stream' });
                                } else {
                                    new Notice('暂不支持该格式导出');
                                    return;
                                }
                                const a = document.createElement('a');
                                a.href = URL.createObjectURL(blob);
                                a.download = finalName;
                                document.body.appendChild(a);
                                a.click();
                                setTimeout(() => {
                                    URL.revokeObjectURL(a.href);
                                    a.remove();
                                }, 100);
                                new Notice('导出成功: ' + finalName);
                            } catch (e) {
                                new Notice('导出失败: ' + (e?.message || e));
                            } finally {
                                if (loading) loading.remove();
                            }
                        });
                        modal.open();
                    });
                }},
            ];
            options.forEach(opt => {
                const item = document.createElement('div');
                item.textContent = opt.label;
                item.style.padding = '12px 24px';
                item.style.cursor = 'pointer';
                item.style.background = '#fff'; // 默认白底
                item.style.color = '#222';      // 默认黑字
                item.onmouseenter = () => {
                    item.style.background = '#e3eafc'; // 悬停浅蓝底
                    item.style.color = '#1765c1';      // 悬停蓝字
                };
                item.onmouseleave = () => {
                    item.style.background = '#fff';
                    item.style.color = '#222';
                };
                item.onclick = (event) => {
                    event.stopPropagation();
                    if (dropdown) dropdown.remove();
                    if (typeof opt.action === 'function') opt.action();
                };
                dropdown!.appendChild(item);
            });
            translateBtn.appendChild(dropdown);
            setTimeout(() => {
                const close = (ev: MouseEvent) => {
                    if (!dropdown) return;
                    if (!dropdown.contains(ev.target as Node) && ev.target !== translateBtn) {
                        dropdown.remove();
                        dropdown = null;
                        document.removeEventListener('mousedown', close);
                    }
                };
                document.addEventListener('mousedown', close);
            }, 10);
        };

        bar.appendChild(pagDiv);
        bar.appendChild(translateBtn);
        this.tablePreviewEl.appendChild(bar);
    }
    private async exportTableFile(format: string) {
        const data = this.showTranslated && this.translatedTableData ? this.translatedTableData : this.tableData;
        if (!data) {
            new Notice('无表格数据，无法导出');
            return;
        }
        // 1. 获取原始文件名或让用户输入
        let defaultName = this._originFileName || '表格名';
        new FileNamePromptModal(this.app, defaultName + '_translate', (fileName) => {
            if (!fileName) return;
            let ext = format;
            if (fileName.endsWith('.' + ext)) fileName = fileName.slice(0, -1 - ext.length);
            let finalName = fileName + '.' + ext;
            const headers = data.headers;
            const rows = data.rows.map(row => [...row]);
            let blob: Blob;
            if (format === 'csv' || format === 'tsv') {
                const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                const csv = XLSX.utils.sheet_to_csv(ws, { FS: format === 'tsv' ? '\t' : ',' });
                blob = new Blob([csv], { type: 'text/' + format });
            } else if (format === 'md') {
                const md = markdownTable([headers, ...rows]);
                blob = new Blob([md], { type: 'text/markdown' });
            } else if (format === 'xlsx' || format === 'ods') {
                const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                const wbout = XLSX.write(wb, { bookType: format, type: 'array' });
                blob = new Blob([wbout], { type: 'application/octet-stream' });
            } else {
                new Notice('暂不支持该格式导出');
                return;
            }
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = finalName;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                URL.revokeObjectURL(a.href);
                a.remove();
            }, 100);
            new Notice('导出成功: ' + finalName);
        }).open();
    }


    private async translateCell(cell: string): Promise<string> {
        const settings = this.plugin.settings;
        if (!cell || !(cell ?? '').toString().trim()) return cell;
        // 判断目标语言
        let targetLang = settings.textTargetLang;
        if (this.targetLangMode === 'auto') {
            if (settings.isBiDirection) {
                targetLang = /[\u4e00-\u9fa5]/.test(cell) ? 'en' : (targetLang === 'zh' ? 'zh-CHS' : targetLang);
            }
        } else if (this.targetLangMode === 'zh') {
            targetLang = this.modelMode === 'composite' ? 'zh-Hans' : 'zh-CHS';
        } else if (this.targetLangMode === 'en') {
            targetLang = 'en';
        }
        // 新增：如果目标语言为中文，且内容包含繁体，则用 t2s 转为简体
        if ((targetLang === 'zh-CHS' || targetLang === 'zh-Hans')) {
            const simplified = t2s(cell);
            if (simplified !== cell) {
                return simplified;
            }
            // 原有逻辑：如果内容全是简体中文，直接跳过
            if (/^[\u4e00-\u9fa5\s，。！？、；："'（）【】《》…—·]*$/.test(cell)) {
                return cell;
            }
        }
        if (targetLang === 'en' && /^[a-zA-Z0-9\s.,!?;:'"()\[\]{}<>@#$%^&*_+=|\\/-]*$/.test(cell)) {
            return cell;
        }
        // 选择翻译模型
        if (this.modelMode === 'composite') {
            // 微软翻译
            const msAdapter = new MicrosoftAdapter(
                settings.microsoftKey || '',
                settings.microsoftRegion || '',
                settings.microsoftEndpoint || '',
                this.app
            );
            const from = /[\u4e00-\u9fa5]/.test(cell) ? 'zh-Hans' : 'en';
            const tLang = targetLang === 'zh' ? 'zh-Hans' : targetLang;
            const translated = await msAdapter.translateText(cell, from, tLang);
            console.log('[翻译日志][MicrosoftAdapter] 原文:', cell, 'from:', from, 'to:', tLang, '返回:', translated);
            return translated || cell;
        } else {
            // 有道翻译
            const translator: TranslateService = new YoudaoAdapter(
                settings.textAppKey,
                settings.textAppSecret,
                this.app,
                settings.serverPort
            );
            const translated = await translator.translateText(cell, 'auto', targetLang);
            console.log('[翻译日志][YoudaoAdapter] 原文:', cell, 'from:auto', 'to:', targetLang, '返回:', translated);
            return translated || cell;
        }
    }
}

// 下拉菜单辅助函数，支持回调
function showDropdown(anchor: HTMLElement, options: string[], onSelect?: (val: string) => void) {
    // 先移除已存在的下拉
    anchor.querySelectorAll('.custom-dropdown').forEach(el => el.remove());
    const dropdown = document.createElement('div');
    dropdown.className = 'custom-dropdown';
    dropdown.style.position = 'absolute';
    dropdown.style.top = '100%';
    dropdown.style.left = '0';
    dropdown.style.background = '#fff';
    dropdown.style.border = '1px solid #ccc';
    dropdown.style.zIndex = '10002'; // 比主菜单更高
    dropdown.style.minWidth = anchor.offsetWidth + 'px';
    dropdown.style.boxShadow = '0 2px 8px #0002';
    dropdown.style.borderRadius = '8px';
    dropdown.style.maxHeight = '60vh';
    dropdown.style.overflowY = 'auto';
    options.forEach(opt => {
        const item = document.createElement('div');
        item.textContent = opt;
        item.style.padding = '12px 24px';
        item.style.cursor = 'pointer';
        item.style.background = '#fff';
        item.style.color = '#222';
        item.onmouseenter = () => { item.style.background = '#e3eafc'; item.style.color = '#1765c1'; };
        item.onmouseleave = () => { item.style.background = '#fff'; item.style.color = '#222'; };
        item.onclick = (event) => {
            event.stopPropagation();
            if (dropdown) dropdown.remove();
            if (typeof onSelect === 'function') onSelect(opt);
        };
        dropdown.appendChild(item);
    });
    anchor.style.position = 'relative'; // 保证anchor为定位父级
    anchor.appendChild(dropdown);
    // 点击其他地方关闭
    setTimeout(() => {
        const close = (ev: MouseEvent) => {
            if (!dropdown.contains(ev.target as Node)) {
                dropdown.remove();
                document.removeEventListener('mousedown', close);
            }
        };
        document.addEventListener('mousedown', close);
    }, 10);
}

// 修正 loadSqlJs 只返回 window.initSqlJs
async function loadSqlJs() {
    if ((window as any).initSqlJs) return (window as any).initSqlJs;
    return new Promise((resolve, reject) => {
        const script = document.createElement('script');
        script.src = 'https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/sql-wasm.js';
        script.onload = () => resolve((window as any).initSqlJs);
        script.onerror = reject;
        document.head.appendChild(script);
    });
}


class FileNamePromptModal extends Modal {
    private defaultName: string;
    private onSubmit: (fileName: string) => void;
    constructor(app: App, defaultName: string, onSubmit: (fileName: string) => void) {
        super(app);
        this.defaultName = defaultName;
        this.onSubmit = onSubmit;
    }
    onOpen() {
        const { contentEl } = this;
        contentEl.createEl('h3', { text: '请输入导出文件名（不含扩展名）' });
        const input = contentEl.createEl('input', { type: 'text', value: this.defaultName });
        input.style.width = '100%';
        input.focus();
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') {
                this.onSubmit(input.value);
                this.close();
            }
        });
        const btn = contentEl.createEl('button', { text: '确定' });
        btn.onclick = () => {
            this.onSubmit(input.value);
            this.close();
        };
    }
}

class TableSqliteOpModal extends Modal {
    plugin: any;
    constructor(app: App, plugin: any) {
        super(app);
        this.plugin = plugin;
    }
    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        // 返回按钮
        const backBtn = contentEl.createEl('button', { text: '返回' });
        backBtn.style.position = 'absolute';
        backBtn.style.left = '16px';
        backBtn.style.top = '16px';
        backBtn.onclick = () => {
            this.close();
            new OtherFunctionModal(this.app, this.plugin).open();
        };
        contentEl.createEl('h2', { text: '表格文件和sqlite3操作' });
        // 主按钮
        const importBtn = contentEl.createEl('button', { text: '把表格文件导入到指定sqlite3' });
        importBtn.style.display = 'block';
        importBtn.style.margin = '40px auto';
        importBtn.onclick = () => {
            this.close();
            new TableToSqliteImportModal(this.app, this.plugin).open();
        };
        // 新增：导出按钮
        const exportBtn = contentEl.createEl('button', { text: '从sqlite3导出指定表格文件' });
        exportBtn.style.display = 'block';
        exportBtn.style.margin = '20px auto';
        exportBtn.onclick = () => {
            this.close();
            new TableFromSqliteExportModal(this.app, this.plugin).open();
        };

    }
}

class TableFromSqliteExportModal extends Modal {
    plugin: any;
    private sqliteFile: File | null = null;
    private sqliteFilePath: string = '';
    private tableName: string = '';
    private exportType: string = '';
    private pickTableBtn: HTMLButtonElement | null = null;
    private exportTypeBtn: HTMLButtonElement | null = null;
    private exportBtn: HTMLButtonElement | null = null;
    private tablePathDiv: HTMLDivElement | null = null;
    private typePathDiv: HTMLDivElement | null = null;
    private tableNames: string[] = [];
    constructor(app: App, plugin: any) {
        super(app);
        this.plugin = plugin;
    }
    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        // 返回按钮
        const backBtn = contentEl.createEl('button', { text: '返回' });
        backBtn.style.position = 'absolute';
        backBtn.style.left = '16px';
        backBtn.style.top = '16px';
        backBtn.onclick = () => {
            this.close();
            new TableSqliteOpModal(this.app, this.plugin).open();
        };
        contentEl.createEl('h2', { text: '表格导出sqlite3' });
        // 选中sqlite3表格按钮
        this.pickTableBtn = contentEl.createEl('button', { text: '选中指定sqlite3的指定表格' });
        this.pickTableBtn.style.display = 'block';
        this.pickTableBtn.style.margin = '40px auto 16px auto';
        this.pickTableBtn.onclick = () => this.pickSqliteAndTable();
        this.tablePathDiv = contentEl.createEl('div');
        this.tablePathDiv.style.textAlign = 'center';
        this.tablePathDiv.style.color = '#888';
        this.tablePathDiv.style.marginBottom = '8px';
        // 导出类型按钮
        this.exportTypeBtn = contentEl.createEl('button', { text: '导出类型' });
        this.exportTypeBtn.style.display = 'block';
        this.exportTypeBtn.style.margin = '16px auto';
        this.exportTypeBtn.onclick = () => this.pickExportType();
        this.typePathDiv = contentEl.createEl('div');
        this.typePathDiv.style.textAlign = 'center';
        this.typePathDiv.style.color = '#888';
        this.typePathDiv.style.marginBottom = '8px';
        // 导出按钮
        this.exportBtn = contentEl.createEl('button', { text: '导出' });
        this.exportBtn.style.display = 'block';
        this.exportBtn.style.margin = '32px auto';
        this.exportBtn.disabled = true;
        this.exportBtn.onclick = () => this.handleExport();
        this.updateExportBtnState();
    }
    private updateExportBtnState() {
        if (this.tableName && this.exportType) {
            this.exportBtn && (this.exportBtn.disabled = false);
        } else {
            this.exportBtn && (this.exportBtn.disabled = true);
        }
    }
    private pickSqliteAndTable() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.sqlite3';
        input.onchange = async () => {
            const file = input.files?.[0];
            if (!file) return;
            this.sqliteFile = file;
            this.sqliteFilePath = file.name;
            // 读取表名
            const arrayBuffer = await file.arrayBuffer();
            const initSqlJs = await loadSqlJs();
            const SQL = await initSqlJs({ locateFile: (file: string) => 'https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/' + file });
            const db = new SQL.Database(new Uint8Array(arrayBuffer));
            const res = db.exec('SELECT name FROM sqlite_master WHERE type="table"');
            this.tableNames = res && res[0] ? res[0].values.flat() : [];
            if (this.tableNames.length === 0) {
                new Notice('该sqlite3文件没有表');
                return;
            }
            // 下拉选择表名
            showDropdown(this.pickTableBtn!, this.tableNames, (val) => {
                this.tableName = val;
                this.tablePathDiv!.textContent = val;
                this.updateExportBtnState();
            });
            // 显示文件名（样式同导入界面）
            this.tablePathDiv!.textContent = file.name;
        };
        input.click();
    }
    private pickExportType() {
        const types = ['csv', 'tsv', 'md', 'xlsx', 'ods'];
        showDropdown(this.exportTypeBtn!, types.map(t => '.' + t), (val) => {
            this.exportType = val.replace('.', '');
            this.typePathDiv!.textContent = val;
            this.updateExportBtnState();
        });
    }
    private async handleExport() {
        if (!this.sqliteFile || !this.tableName || !this.exportType) return;
        const defaultName = this.tableName;
        new FileNamePromptModal(this.app, defaultName, async (fileName) => {
            if (!fileName) return;
            let ext = this.exportType;
            if (fileName.endsWith('.' + ext)) fileName = fileName.slice(0, -1 - ext.length);
            let finalName = fileName + '.' + ext;
            // 读取sqlite3表数据
            if (!this.sqliteFile) return;
            const arrayBuffer = await this.sqliteFile.arrayBuffer();
            const initSqlJs = await loadSqlJs();
            const SQL = await initSqlJs({ locateFile: (file: string) => 'https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/' + file });
            const db = new SQL.Database(new Uint8Array(arrayBuffer));
            const res = db.exec(`SELECT * FROM "${this.tableName}"`);
            if (!res || !res[0]) {
                new Notice('表格数据为空或读取失败');
                return;
            }
            const headers = res[0].columns;
            const rows = res[0].values;
            let blob: Blob;
            if (ext === 'csv' || ext === 'tsv') {
                const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                const csv = XLSX.utils.sheet_to_csv(ws, { FS: ext === 'tsv' ? '\t' : ',' });
                blob = new Blob([csv], { type: 'text/' + ext });
            } else if (ext === 'md') {
                const md = markdownTable([headers, ...rows]);
                blob = new Blob([md], { type: 'text/markdown' });
            } else if (ext === 'xlsx' || ext === 'ods') {
                const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
                const wb = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
                const wbout = XLSX.write(wb, { bookType: ext, type: 'array' });
                blob = new Blob([wbout], { type: 'application/octet-stream' });
            } else {
                new Notice('暂不支持该格式导出');
                return;
            }
            const a = document.createElement('a');
            a.href = URL.createObjectURL(blob);
            a.download = finalName;
            document.body.appendChild(a);
            a.click();
            setTimeout(() => {
                URL.revokeObjectURL(a.href);
                a.remove();
            }, 1000);
        }).open();
    }
}

class TableToSqliteImportModal extends Modal {
    plugin: any;
    private tableFile: File | null = null;
    private tableFilePath: string = '';
    private sqliteFile: File | null = null;
    private sqliteFilePath: string = '';
    private importBtn: HTMLButtonElement | null = null;
    private pickTableBtn: HTMLButtonElement | null = null;
    private pickSqliteBtn: HTMLButtonElement | null = null;
    private tablePathDiv: HTMLDivElement | null = null;
    private sqlitePathDiv: HTMLDivElement | null = null;
    constructor(app: App, plugin: any) {
        super(app);
        this.plugin = plugin;
    }
    onOpen() {
        const { contentEl } = this;
        contentEl.empty();
        // 返回按钮
        const backBtn = contentEl.createEl('button', { text: '返回' });
        backBtn.style.position = 'absolute';
        backBtn.style.left = '16px';
        backBtn.style.top = '16px';
        backBtn.onclick = () => {
            this.close();
            new TableSqliteOpModal(this.app, this.plugin).open();
        };
        contentEl.createEl('h2', { text: '表格导入sqlite3' });
        // 选中表格文件按钮
        this.pickTableBtn = contentEl.createEl('button', { text: '选中指定表格文件' });
        this.pickTableBtn.style.display = 'block';
        this.pickTableBtn.style.margin = '40px auto 16px auto';
        this.pickTableBtn.onclick = () => this.pickTableFile();
        this.tablePathDiv = contentEl.createEl('div');
        this.tablePathDiv.style.textAlign = 'center';
        this.tablePathDiv.style.color = '#888';
        // 选中sqlite3按钮
        this.pickSqliteBtn = contentEl.createEl('button', { text: '选中指定sqlite3' });
        this.pickSqliteBtn.style.display = 'block';
        this.pickSqliteBtn.style.margin = '16px auto';
        this.pickSqliteBtn.onclick = () => this.pickSqliteFile();
        this.sqlitePathDiv = contentEl.createEl('div');
        this.sqlitePathDiv.style.textAlign = 'center';
        this.sqlitePathDiv.style.color = '#888';
        // 导入按钮
        this.importBtn = contentEl.createEl('button', { text: '导入' });
        this.importBtn.style.display = 'block';
        this.importBtn.style.margin = '32px auto';
        this.importBtn.disabled = true;
        this.importBtn.onclick = () => this.handleImport();
        this.updateImportBtnState();
    }
    updateImportBtnState() {
        if (this.importBtn) {
            this.importBtn.disabled = !(this.tableFile && this.sqliteFile);
        }
    }
    pickTableFile() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.csv,.tsv,.md,.xlsx,.ods';
        input.onchange = () => {
            const file = input.files?.[0];
            if (!file) return;
            this.tableFile = file;
            this.tableFilePath = file.name;
            if (this.tablePathDiv) this.tablePathDiv.textContent = file.name;
            this.updateImportBtnState();
        };
        input.click();
    }
    pickSqliteFile() {
        const input = document.createElement('input');
        input.type = 'file';
        input.accept = '.sqlite3';
        input.onchange = () => {
            const file = input.files?.[0];
            if (!file) return;
            this.sqliteFile = file;
            this.sqliteFilePath = file.name;
            if (this.sqlitePathDiv) this.sqlitePathDiv.textContent = file.name;
            this.updateImportBtnState();
        };
        input.click();
    }
    async handleImport() {
        if (!this.tableFile || !this.sqliteFile) return;
        // 1. 读取表格文件内容
        let tableData: { headers: string[], rows: string[][] } | null = null;
        let tableName = this.tableFile.name.replace(/\.[^.]+$/, '');
        let ext = this.tableFile.name.split('.').pop()?.toLowerCase();
        let sheetName = '';
        if (ext === 'xlsx' || ext === 'ods') {
            const arrayBuffer = await this.tableFile.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            if (workbook.SheetNames.length > 1) {
                // 多sheet，弹窗让用户选择
                await new Promise<void>(resolve => {
                    const modal = new Modal(this.app);
                    modal.contentEl.createEl('h3', { text: '请选择要导入的sheet' });
                    workbook.SheetNames.forEach(name => {
                        const btn = modal.contentEl.createEl('button', { text: name });
                        btn.style.margin = '8px';
                        btn.onclick = () => {
                            sheetName = name;
                            modal.close();
                            resolve();
                        };
                    });
                    modal.open();
                });
            } else {
                sheetName = workbook.SheetNames[0];
            }
            const worksheet = workbook.Sheets[sheetName];
            const ref = worksheet['!ref'];
            if (!ref) {
                new Notice('该表为空或无数据');
                return;
            }
            const range = XLSX.utils.decode_range(ref);
            const maxRow = range.e.r;
            const maxCol = range.e.c;
            const rows: string[][] = [];
            for (let r = 0; r <= maxRow; r++) {
                const row: string[] = [];
                for (let c = 0; c <= maxCol; c++) {
                    const cellAddress = XLSX.utils.encode_cell({ c, r });
                    const cell = worksheet[cellAddress];
                    row.push(cell ? String(cell.v) : '');
                }
                rows.push(row);
            }
            const headers = rows.length > 0 ? rows[0] : [];
            const dataRows = rows.slice(1);
            tableData = { headers, rows: dataRows };
        } else {
            const text = await this.tableFile.text();
            if (ext === 'tsv') {
                tableData = parseCsvTsv(text, '\t');
            } else if (ext === 'md') {
                tableData = parseMarkdownTable(text);
            } else {
                tableData = parseCsvTsv(text, ',');
            }
        }
        if (!tableData) {
            new Notice('表格解析失败');
            return;
        }
        // 2. 检查表头字段名
        const headers = tableData.headers;
        const headerSet = new Set<string>();
        let illegalField = '';
        let duplicateField = '';
        for (const h of headers) {
            if (!/^[_a-zA-Z][_a-zA-Z0-9]*$/.test(h)) {
                illegalField = h;
                break;
            }
            if (headerSet.has(h)) {
                duplicateField = h;
                break;
            }
            headerSet.add(h);
        }
        if (illegalField) {
            this.showErrorModal(`字段名"${illegalField}"不合法，字段名只能以字母或下划线开头，只能包含字母、数字和下划线。请修改后重试。`);
            return;
        }
        if (duplicateField) {
            this.showErrorModal(`字段名"${duplicateField}"重复，请修改后重试。`);
            return;
        }
        // 3. 读取sqlite3文件
        const arrayBuffer = await this.sqliteFile.arrayBuffer();
        const initSqlJs = await loadSqlJs();
        const SQL = await initSqlJs({ locateFile: (file: string) => 'https://cdn.jsdelivr.net/npm/sql.js@1.8.0/dist/' + file });
        const db = new SQL.Database(new Uint8Array(arrayBuffer));
        // 4. 检查表名是否存在
        let tableExists = false;
        try {
            const res = db.exec(`SELECT name FROM sqlite_master WHERE type='table' AND name='${tableName}'`);
            if (res && res[0] && res[0].values.length > 0) tableExists = true;
        } catch {}
        if (tableExists) {
            // 弹窗提示是否覆盖
            const shouldContinue = await new Promise<boolean>((resolve) => {
                const modal = new Modal(this.app);
                modal.contentEl.createEl('h3', { text: `表"${tableName}"已存在，是否覆盖？` });
                const coverBtn = modal.contentEl.createEl('button', { text: '覆盖' });
                coverBtn.style.margin = '8px';
                coverBtn.onclick = () => { modal.close(); resolve(true); };
                const cancelBtn = modal.contentEl.createEl('button', { text: '取消' });
                cancelBtn.style.margin = '8px';
                cancelBtn.onclick = () => { modal.close(); resolve(false); };
                modal.open();
            });
            if (!shouldContinue) return; // 用户点取消，直接终止整个导入
            // 删除原表
            db.exec(`DROP TABLE IF EXISTS "${tableName}"`);
        }
        // 5. 新建表
        const colDefs = headers.map(h => `"${h}" TEXT`).join(', ');
        db.exec(`CREATE TABLE "${tableName}" (${colDefs})`);
        // 6. 逐行插入数据，进度条
        let imported = 0;
        const total = tableData.rows.length;
        let stop = false;
        let progressModal: Modal | null = null;
        let progressBar: HTMLDivElement | null = null;
        let progressText: HTMLDivElement | null = null;
        const showProgress = () => {
            if (!progressModal) {
                progressModal = new Modal(this.app);
                progressModal.contentEl.createEl('h3', { text: '正在导入...' });
                progressBar = progressModal.contentEl.createEl('div');
                progressBar.style.height = '16px';
                progressBar.style.background = '#eee';
                progressBar.style.borderRadius = '8px';
                progressBar.style.margin = '16px 0';
                progressBar.style.position = 'relative';
                progressText = progressModal.contentEl.createEl('div');
                progressText.style.textAlign = 'center';
                progressText.style.margin = '8px 0';
                progressModal.open();
            }
            if (progressBar && progressText) {
                const percent = total ? Math.round(imported / total * 100) : 0;
                progressBar.innerHTML = `<div style="height:100%;width:${percent}%;background:#1a73e8;border-radius:8px;"></div>`;
                progressText.textContent = `${imported}/${total}`;
            }
        };
        showProgress();
        for (let i = 0; i < total; i++) {
            if (stop) break;
            const row = tableData.rows[i];
            try {
                const placeholders = headers.map(() => '?').join(',');
                db.run(`INSERT INTO "${tableName}" (${headers.map(h => '"'+h+'"').join(',')}) VALUES (${placeholders})`, row);
                imported++;
                showProgress(); 
            } catch (e) {
                await new Promise<void>((resolve) => {
                    const modal = new Modal(this.app);
                    modal.contentEl.createEl('h3', { text: `第${i+2}行导入失败` });
                    modal.contentEl.createEl('div', { text: `原因: ${e.message}` });
                    modal.contentEl.createEl('pre', { text: JSON.stringify(row) });
                    const skipBtn = modal.contentEl.createEl('button', { text: '跳过该行继续导入' });
                    skipBtn.style.margin = '8px';
                    skipBtn.onclick = () => { modal.close(); resolve(); };
                    const stopBtn = modal.contentEl.createEl('button', { text: '终止导入' });
                    stopBtn.style.margin = '8px';
                    stopBtn.onclick = () => { stop = true; modal.close(); resolve(); };
                    modal.open();
                });
            }
        }
        if (progressModal !== null) (progressModal as Modal).close();
        // 7. 导入完成提示
        const doneModal = new Modal(this.app);
        doneModal.contentEl.createEl('h3', { text: '导入完成' });
        doneModal.contentEl.createEl('div', { text: `共导入${imported}行` });
        doneModal.contentEl.createEl('div', { text: `表名: ${tableName}` });
        doneModal.contentEl.createEl('div', { text: `字段数: ${headers.length}` });
        const okBtn = doneModal.contentEl.createEl('button', { text: '确定' });
        okBtn.onclick = () => doneModal.close();
        doneModal.open();
        // 8. 导出sqlite3文件
        const data = db.export();
        const blob = new Blob([data], { type: 'application/octet-stream' });
        const a = document.createElement('a');
        a.href = URL.createObjectURL(blob);
        a.download = this.sqliteFile.name;
        a.click();
        setTimeout(() => URL.revokeObjectURL(a.href), 1000);
    }
    showErrorModal(msg: string) {
        const modal = new Modal(this.app);
        modal.contentEl.createEl('h3', { text: '导入失败' });
        modal.contentEl.createEl('div', { text: msg });
        const btn = modal.contentEl.createEl('button', { text: '确定' });
        btn.onclick = () => modal.close();
        modal.open();
    }
}

