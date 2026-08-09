/**
 * Auto Retouch — Photoshop 2024/2025 UXP Plugin (v3 - evalScript 版)
 * 能识别、面板正常显示，但执行时 evalScript 不存在会报错
 */

const { action: psAction } = require("photoshop");

// ──────────────────────────────────────────────────────────────
// ExtendScript Core (embedded as string)
// ──────────────────────────────────────────────────────────────
const JSX_CORE = `
"use strict";

function _checkDoc() {
    return app.documents.length > 0;
}

function _gaussianBlur(radius) {
    var d = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
    d.putReference(charIDToTypeID("null"), ref);
    d.putUnitDouble(charIDToTypeID("Rds "), charIDToTypeID("Pxl "), radius);
    executeAction(charIDToTypeID("GsnB"), d, DialogModes.NO);
}

function _lensBlur(radius, shape, blades, rotation, depthChannel, invertDepth,
                    highlightGain, highlightThresh, noiseAmount) {
    var d = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
    d.putReference(charIDToTypeID("null"), ref);
    d.putUnitDouble(charIDToTypeID("Rds "), charIDToTypeID("Pxl "), radius);
    var shapeIDs = [
        charIDToTypeID("shCr"), charIDToTypeID("shTr"), charIDToTypeID("shSq"),
        charIDToTypeID("shPg"), charIDToTypeID("shHg"), charIDToTypeID("shOg")
    ];
    d.putEnumerated(charIDToTypeID("shAp"), charIDToTypeID("shAp"), shapeIDs[shape] || shapeIDs[0]);
    d.putInteger(charIDToTypeID("blnd"), blades);
    d.putUnitDouble(charIDToTypeID("rotd"), charIDToTypeID("Ang "), rotation);
    if (depthChannel) {
        d.putEnumerated(charIDToTypeID("blAd"), charIDToTypeID("blAd"), charIDToTypeID("blCh"));
        d.putReference(charIDToTypeID("chAn"),
            new ActionReference().putName(charIDToTypeID("Chnl"), depthChannel.name));
        d.putBoolean(charIDToTypeID("inDe"), invertDepth);
    } else {
        d.putEnumerated(charIDToTypeID("blAd"), charIDToTypeID("blAd"), charIDToTypeID("blNn"));
    }
    d.putInteger(charIDToTypeID("hoTi"), highlightGain);
    d.putInteger(charIDToTypeID("hoSt"), highlightThresh);
    d.putUnitDouble(charIDToTypeID("noAm"), charIDToTypeID("Pxl "), noiseAmount);
    d.putEnumerated(charIDToTypeID("noDi"), charIDToTypeID("noDi"), charIDToTypeID("noUn"));
    d.putBoolean(charIDToTypeID("grMo"), false);
    executeAction(charIDToTypeID("LnsB"), d, DialogModes.NO);
}

function _buildDepthFromSubject(feather) {
    var doc = app.activeDocument;
    var ch = doc.channels.add();
    ch.name = "__depth_map__";
    ch.selection.store(doc.selection);
    doc.selection.load(ch.selection);
    _featherSelection(Math.max(feather * 2, 20));
    doc.selection.store(ch.selection);
    _deselectAll();
    doc.activeChannel = ch;
    _gaussianBlur(Math.max(feather, 15));
    doc.activeChannel = doc.channels[0];
    return ch;
}

function _invertSelection() {
    var d = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
    d.putReference(charIDToTypeID("null"), ref);
    executeAction(charIDToTypeID("Invr"), d, DialogModes.NO);
}

function _deselectAll() {
    var d = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
    d.putReference(charIDToTypeID("null"), ref);
    d.putEnumerated(charIDToTypeID("T   "), charIDToTypeID("Ordn"), charIDToTypeID("None"));
    executeAction(charIDToTypeID("Dplc"), d, DialogModes.NO);
}

function _featherSelection(radius) {
    var d = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
    d.putReference(charIDToTypeID("null"), ref);
    d.putUnitDouble(charIDToTypeID("Rds "), charIDToTypeID("Pxl "), radius);
    executeAction(charIDToTypeID("Fthr"), d, DialogModes.NO);
}

function _addMaskFromSelection(invert) {
    var d = new ActionDescriptor();
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("mask"));
    ref.putEnumerated(charIDToTypeID("Chnl"), charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
    d.putReference(charIDToTypeID("null"), ref);
    if (invert) d.putBoolean(charIDToTypeID("Invr"), true);
    executeAction(charIDToTypeID("Dplc"), d, DialogModes.NO);
}

function _stampVisible(name) {
    var doc = app.activeDocument;
    var l = doc.artLayers.add();
    l.name = name;
    var ref = new ActionReference();
    ref.putProperty(charIDToTypeID("Lyr "), charIDToTypeID("fStV"));
    var d = new ActionDescriptor();
    d.putReference(charIDToTypeID("null"), ref);
    executeAction(charIDToTypeID("Dplc"), d, DialogModes.NO);
    return l;
}

function _buildSkinProtectMask(strength) {
    var doc = app.activeDocument;
    var orig = doc.activeLayer;
    var dl = orig.duplicate();
    dl.name = "__detail_detect__";
    var d = new ActionDescriptor();
    d.putUnitDouble(charIDToTypeID("Rds "), charIDToTypeID("Pxl "), Math.max(3, strength * 2));
    executeAction(charIDToTypeID("HghP"), d, DialogModes.NO);
    dl.invert();
    doc.selection.load(dl.selection);
    _invertSelection();
    var ch = doc.channels.add();
    ch.name = "__skin_mask__";
    ch.selection.store(doc.selection);
    dl.remove();
    _deselectAll();
    return ch;
}

function _selectSubject() {
    try {
        var d = new ActionDescriptor();
        var ref = new ActionReference();
        ref.putProperty(charIDToTypeID("Chnl"), charIDToTypeID("fsel"));
        d.putReference(charIDToTypeID("null"), ref);
        d.putBoolean(charIDToTypeID("AntA"), true);
        d.putBoolean(charIDToTypeID("Cntg"), false);
        executeAction(stringIDToTypeID("selectSubject"), d, DialogModes.NO);
        return true;
    } catch (e) {
        try {
            app.activeDocument.selection.selectAll();
            return false;
        } catch (e2) {
            return false;
        }
    }
}

function autoSkinRetouch(strength, detail, faceOnly) {
    if (!_checkDoc()) return "ERROR_no_document";
    try {
        var doc = app.activeDocument;
        var orig = doc.activeLayer;
        app.displayDialogs = DialogModes.NO;
        strength = Math.max(1, Math.min(20, strength));
        detail = Math.max(0, Math.min(100, detail));
        var skm = null;
        if (faceOnly) skm = _buildSkinProtectMask(strength);
        var bl = orig.duplicate();
        bl.name = "Retouch_Skin_Blur";
        doc.activeLayer = bl;
        _gaussianBlur(strength);
        if (faceOnly && skm) {
            doc.selection.load(skm.selection);
            _addMaskFromSelection(false);
            _deselectAll();
        }
        bl.blendMode = BlendMode.LIGHTEN;
        bl.opacity = Math.round(100 * (1 - detail / 100));
        var rs = _stampVisible("Retouch_Skin_Result");
        rs.opacity = 100;
        bl.remove();
        try { skm.remove(); } catch (e) {}
        doc.activeLayer = rs;
        return "OK";
    } catch (e) {
        return "ERROR_" + e.message;
    }
}

function autoBackgroundBlur(strength, feather) {
    if (!_checkDoc()) return "ERROR_no_document";
    try {
        var doc = app.activeDocument;
        var orig = doc.activeLayer;
        app.displayDialogs = DialogModes.NO;
        strength = Math.max(3, Math.min(60, strength));
        feather = Math.max(0, Math.min(200, feather));
        _selectSubject();
        var fg = doc.channels.add();
        fg.name = "__fg_mask__";
        fg.selection.store(doc.selection);
        _invertSelection();
        var dep = _buildDepthFromSubject(feather);
        var bl = orig.duplicate();
        bl.name = "Retouch_BG_LensBlur";
        doc.activeLayer = bl;
        _deselectAll();
        _lensBlur(strength, 4, 6, 0, dep, true, 30, 250, 1.5);
        doc.selection.load(fg.selection);
        _addMaskFromSelection(true);
        _deselectAll();
        try { dep.remove(); } catch (e) {}
        try { fg.remove(); } catch (e) {}
        doc.activeLayer = bl;
        return "OK";
    } catch (e) {
        return "ERROR_" + e.message;
    }
}
`;

// ──────────────────────────────────────────────────────────────
// evalScript 调用 v3 版本
// ──────────────────────────────────────────────────────────────
let coreLoaded = false;

async function ensureCore() {
    if (coreLoaded) return;
    await psAction.evalScript(JSX_CORE);
    coreLoaded = true;
}

async function runFn(name, args) {
    await ensureCore();
    const argStr = args.map((a) => {
        if (typeof a === "boolean") return a ? "true" : "false";
        if (typeof a === "number") return String(a);
        return '"' + String(a).replace(/"/g, '\\"') + '"';
    }).join(",");
    const script = name + "(" + argStr + ")";
    return await psAction.evalScript(script);
}

// ──────────────────────────────────────────────────────────────
// UI 绑定
// ──────────────────────────────────────────────────────────────
const els = {
    skinStrength: document.getElementById("skinStrength"),
    skinStrengthVal: document.getElementById("skinStrengthVal"),
    skinDetail: document.getElementById("skinDetail"),
    skinDetailVal: document.getElementById("skinDetailVal"),
    skinFaceOnly: document.getElementById("skinFaceOnly"),
    btnSkin: document.getElementById("btnSkin"),

    blurStrength: document.getElementById("blurStrength"),
    blurStrengthVal: document.getElementById("blurStrengthVal"),
    blurFeather: document.getElementById("blurFeather"),
    blurFeatherVal: document.getElementById("blurFeatherVal"),
    btnBlur: document.getElementById("btnBlur"),

    quickSkin: document.getElementById("quickSkin"),
    quickBlur: document.getElementById("quickBlur"),
    btnQuick: document.getElementById("btnQuick"),

    statusBar: document.getElementById("statusBar"),
    statusText: document.getElementById("statusText"),
};

function setStatus(text, type) {
    els.statusText.textContent = text;
    els.statusBar.classList.remove("working", "error");
    if (type === "working") els.statusBar.classList.add("working");
    if (type === "error") els.statusBar.classList.add("error");
}

function setButtonsEnabled(enabled) {
    els.btnSkin.disabled = !enabled;
    els.btnBlur.disabled = !enabled;
    els.btnQuick.disabled = !enabled;
}

function bindSlider(el, valEl) {
    el.addEventListener("input", () => {
        valEl.textContent = el.value;
    });
}

bindSlider(els.skinStrength, els.skinStrengthVal);
bindSlider(els.skinDetail, els.skinDetailVal);
bindSlider(els.blurStrength, els.blurStrengthVal);
bindSlider(els.blurFeather, els.blurFeatherVal);

// ──────────────────────────────────────────────────────────────
// 按钮事件
// ──────────────────────────────────────────────────────────────
els.btnSkin.addEventListener("click", async () => {
    setButtonsEnabled(false);
    setStatus("🧴 磨皮中...", "working");
    try {
        const r = await runFn("autoSkinRetouch", [
            parseInt(els.skinStrength.value, 10),
            parseInt(els.skinDetail.value, 10),
            els.skinFaceOnly.checked,
        ]);
        const ok = r && r.indexOf("OK") >= 0;
        setStatus(ok ? "✅ 磨皮完成" : "❌ " + (r || "未知错误"), ok ? null : "error");
    } catch (e) {
        setStatus("❌ " + e.message, "error");
    }
    setButtonsEnabled(true);
});

els.btnBlur.addEventListener("click", async () => {
    setButtonsEnabled(false);
    setStatus("📷 虚化中...", "working");
    try {
        const r = await runFn("autoBackgroundBlur", [
            parseInt(els.blurStrength.value, 10),
            parseInt(els.blurFeather.value, 10),
        ]);
        const ok = r && r.indexOf("OK") >= 0;
        setStatus(ok ? "✅ 背景虚化完成" : "❌ " + (r || "未知错误"), ok ? null : "error");
    } catch (e) {
        setStatus("❌ " + e.message, "error");
    }
    setButtonsEnabled(true);
});

els.btnQuick.addEventListener("click", async () => {
    const doSkin = els.quickSkin.checked;
    const doBlur = els.quickBlur.checked;
    if (!doSkin && !doBlur) {
        setStatus("⚠️ 请至少选一项", "error");
        return;
    }
    setButtonsEnabled(false);
    try {
        if (doSkin) {
            setStatus("🧴 磨皮中...", "working");
            const r = await runFn("autoSkinRetouch", [
                parseInt(els.skinStrength.value, 10),
                parseInt(els.skinDetail.value, 10),
                els.skinFaceOnly.checked,
            ]);
            if (!(r && r.indexOf("OK") >= 0)) {
                setStatus("❌ 磨皮失败: " + r, "error");
                setButtonsEnabled(true);
                return;
            }
        }
        if (doBlur) {
            setStatus("📷 虚化中...", "working");
            const r = await runFn("autoBackgroundBlur", [
                parseInt(els.blurStrength.value, 10),
                parseInt(els.blurFeather.value, 10),
            ]);
            if (!(r && r.indexOf("OK") >= 0)) {
                setStatus("❌ 虚化失败: " + r, "error");
                setButtonsEnabled(true);
                return;
            }
        }
        setStatus("✨ 全部完成", null);
    } catch (e) {
        setStatus("❌ " + e.message, "error");
    }
    setButtonsEnabled(true);
});

setStatus("就绪", null);
