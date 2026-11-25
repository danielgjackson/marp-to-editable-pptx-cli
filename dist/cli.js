#!/usr/bin/env node
"use strict";
var __create = Object.create;
var __defProp = Object.defineProperty;
var __getOwnPropDesc = Object.getOwnPropertyDescriptor;
var __getOwnPropNames = Object.getOwnPropertyNames;
var __getProtoOf = Object.getPrototypeOf;
var __hasOwnProp = Object.prototype.hasOwnProperty;
var __export = (target, all) => {
  for (var name in all)
    __defProp(target, name, { get: all[name], enumerable: true });
};
var __copyProps = (to, from, except, desc) => {
  if (from && typeof from === "object" || typeof from === "function") {
    for (let key of __getOwnPropNames(from))
      if (!__hasOwnProp.call(to, key) && key !== except)
        __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });
  }
  return to;
};
var __toESM = (mod, isNodeMode, target) => (target = mod != null ? __create(__getProtoOf(mod)) : {}, __copyProps(
  // If the importer is in node compatibility mode or this is not an ESM
  // file that has been converted to a CommonJS file using a Babel-
  // compatible transform (i.e. "__esModule" has not been set), then set
  // "default" to the CommonJS "module.exports" for node compatibility.
  isNodeMode || !mod || !mod.__esModule ? __defProp(target, "default", { value: mod, enumerable: true }) : target,
  mod
));
var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);

// src/cli.ts
var cli_exports = {};
__export(cli_exports, {
  detectBrowserPath: () => detectBrowserPath
});
module.exports = __toCommonJS(cli_exports);
var import_node_fs3 = require("node:fs");

// node_modules/marp-to-editable-pptx/src/native-pptx/browser.ts
var import_node_fs = require("node:fs");
var import_node_os = require("node:os");
var import_node_path = __toESM(require("node:path"));
var import_browsers = require("@puppeteer/browsers");
function detectBrowserPath(preference = "auto", explicitPath) {
  if (explicitPath && (0, import_node_fs.existsSync)(explicitPath)) return explicitPath;
  const platform = detectPlatform();
  if (!platform) return void 0;
  const candidates = [];
  if (preference === "chrome" || preference === "auto") {
    candidates.push({
      browser: import_browsers.Browser.CHROME,
      channel: import_browsers.ChromeReleaseChannel.STABLE
    });
  }
  for (const { browser, channel } of candidates) {
    try {
      const executablePath = (0, import_browsers.computeSystemExecutablePath)({
        browser,
        platform,
        channel
      });
      if ((0, import_node_fs.existsSync)(executablePath)) return executablePath;
    } catch {
    }
  }
  if (preference === "edge" || preference === "auto") {
    const edgePath = edgeExecutablePaths().find(
      (candidate) => (0, import_node_fs.existsSync)(candidate)
    );
    if (edgePath) return edgePath;
  }
  return void 0;
}
function edgeExecutablePaths() {
  switch (process.platform) {
    case "win32":
      return [
        process.env["PROGRAMFILES(X86)"],
        process.env.PROGRAMFILES,
        process.env.LOCALAPPDATA
      ].filter((root) => !!root).map(
        (root) => import_node_path.default.win32.join(root, "Microsoft", "Edge", "Application", "msedge.exe")
      );
    case "darwin":
      return [
        "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
        import_node_path.default.posix.join(
          (0, import_node_os.homedir)(),
          "Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge"
        )
      ];
    case "linux":
      return ["/usr/bin/microsoft-edge", "/usr/bin/microsoft-edge-stable"];
    default:
      return [];
  }
}
function detectPlatform() {
  switch (process.platform) {
    case "win32":
      return process.arch === "x64" ? import_browsers.BrowserPlatform.WIN64 : import_browsers.BrowserPlatform.WIN32;
    case "darwin":
      return import_browsers.BrowserPlatform.MAC;
    case "linux":
      return import_browsers.BrowserPlatform.LINUX;
    default:
      return void 0;
  }
}

// src/convert.ts
var import_node_crypto = require("node:crypto");
var import_promises2 = require("node:fs/promises");
var import_node_path3 = __toESM(require("node:path"));
var import_node_util2 = require("node:util");
var import_marp_cli = require("@marp-team/marp-cli");

// node_modules/marp-to-editable-pptx/src/native-pptx/index.ts
var import_promises = require("node:fs/promises");
var import_node_perf_hooks = require("node:perf_hooks");
var import_node_url2 = require("node:url");
var import_puppeteer_core = __toESM(require("puppeteer-core"));

// dom-walker-script:dom-walker-script
var DOM_WALKER_SCRIPT = 'var DomWalker = (() => {\n  var __defProp = Object.defineProperty;\n  var __getOwnPropDesc = Object.getOwnPropertyDescriptor;\n  var __getOwnPropNames = Object.getOwnPropertyNames;\n  var __hasOwnProp = Object.prototype.hasOwnProperty;\n  var __export = (target, all) => {\n    for (var name in all)\n      __defProp(target, name, { get: all[name], enumerable: true });\n  };\n  var __copyProps = (to, from, except, desc) => {\n    if (from && typeof from === "object" || typeof from === "function") {\n      for (let key of __getOwnPropNames(from))\n        if (!__hasOwnProp.call(to, key) && key !== except)\n          __defProp(to, key, { get: () => from[key], enumerable: !(desc = __getOwnPropDesc(from, key)) || desc.enumerable });\n    }\n    return to;\n  };\n  var __toCommonJS = (mod) => __copyProps(__defProp({}, "__esModule", { value: true }), mod);\n\n  // node_modules/marp-to-editable-pptx/src/native-pptx/dom-walker.ts\n  var dom_walker_exports = {};\n  __export(dom_walker_exports, {\n    extractSlides: () => extractSlides\n  });\n\n  // node_modules/marp-to-editable-pptx/src/native-pptx/dom-reader.ts\n  function createDomSourceTracker() {\n    const sourceIds = /* @__PURE__ */ new WeakMap();\n    let nextSourceId = 1;\n    return {\n      sourceFor(element) {\n        let id = sourceIds.get(element);\n        if (!id) {\n          id = `node-${nextSourceId++}`;\n          sourceIds.set(element, id);\n        }\n        const parent = element.parentElement;\n        const parentId = parent ? sourceIds.get(parent) : void 0;\n        const segments = [];\n        let current = element;\n        while (current?.parentElement) {\n          const siblings = Array.from(current.parentElement.children);\n          segments.unshift(siblings.indexOf(current));\n          current = current.parentElement;\n        }\n        return {\n          id,\n          domPath: segments.length > 0 ? `/root/${segments.join("/")}` : "/root",\n          ...parentId ? { parentId } : {}\n        };\n      }\n    };\n  }\n  function computeAutoScaleFactor(runs, style, rectHeight, transformScale = 1) {\n    const lineHeight = parseFloat(style.lineHeight) || 0;\n    if (lineHeight <= 0 || rectHeight <= 0) return 1;\n    let lastContentRun = runs.length - 1;\n    while (lastContentRun >= 0 && runs[lastContentRun].breakLine) lastContentRun--;\n    const lineCount = runs.slice(0, lastContentRun + 1).filter((run) => run.breakLine).length + 1;\n    const nonTextHeight = (parseFloat(style.paddingTop) || 0) + (parseFloat(style.paddingBottom) || 0) + (parseFloat(style.borderTopWidth) || 0) + (parseFloat(style.borderBottomWidth) || 0);\n    const contentHeight = rectHeight - nonTextHeight * transformScale;\n    if (contentHeight <= 0) return 1;\n    const naturalHeight = lineCount * lineHeight * transformScale;\n    const fitScale = naturalHeight <= contentHeight * 1.05 ? 1 : contentHeight / naturalHeight;\n    return transformScale * fitScale;\n  }\n  function applyAutoScale(runs, elementFontSize, scaleFactor) {\n    if (Math.abs(scaleFactor - 1) < Number.EPSILON) return elementFontSize;\n    for (const run of runs) {\n      if (!run.breakLine && run.fontSize) run.fontSize *= scaleFactor;\n    }\n    return elementFontSize * scaleFactor;\n  }\n  function parseTransformScale(transform) {\n    if (!transform || transform === "none") return 1;\n    const values = transform.match(/-?\\d*\\.?\\d+(?:e[-+]?\\d+)?/gi)?.map(Number);\n    if (!values) return 1;\n    if (transform.startsWith("scale(")) {\n      const scaleX = values[0] ?? 1;\n      const scaleY = values[1] ?? scaleX;\n      return Math.sqrt(scaleX * scaleY);\n    }\n    if (transform.startsWith("matrix3d") && values.length >= 16) {\n      const scaleX = Math.hypot(values[0], values[1], values[2]);\n      const scaleY = Math.hypot(values[4], values[5], values[6]);\n      return Math.sqrt(scaleX * scaleY);\n    }\n    if (values.length >= 6) {\n      const scaleX = Math.hypot(values[0], values[1]);\n      const scaleY = Math.hypot(values[2], values[3]);\n      return Math.sqrt(scaleX * scaleY);\n    }\n    return 1;\n  }\n  function computeElementTransformScale(element) {\n    const view = element.ownerDocument.defaultView;\n    if (!view) return 1;\n    let scale = 1;\n    let current = element;\n    while (current) {\n      scale *= parseTransformScale(view.getComputedStyle(current).transform);\n      if (current.parentElement) {\n        current = current.parentElement;\n        continue;\n      }\n      const root = current.getRootNode();\n      current = typeof ShadowRoot !== "undefined" && root instanceof ShadowRoot ? root.host : null;\n    }\n    return Number.isFinite(scale) && scale > 0 ? scale : 1;\n  }\n\n  // node_modules/marp-to-editable-pptx/src/native-pptx/dom-walker.ts\n  function extractSlides(root = document) {\n    const { sourceFor: sourceRefFor } = createDomSourceTracker();\n    function findBackgroundColor(section) {\n      const style = getComputedStyle(section);\n      const bg = style.backgroundColor;\n      if (bg && bg !== "transparent" && bg !== "rgba(0, 0, 0, 0)") {\n        return bg;\n      }\n      const bgImage = style.backgroundImage;\n      if (bgImage && bgImage !== "none") {\n        const colorMatches = bgImage.match(\n          /rgba?\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+(?:\\s*,\\s*[\\d.]+)?\\s*\\)/g\n        );\n        if (colorMatches && colorMatches.length > 0) {\n          for (let i = colorMatches.length - 1; i >= 0; i--) {\n            const c = colorMatches[i];\n            if (c !== "rgba(0, 0, 0, 0)") {\n              const alphaMatch = c.match(\n                /rgba\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*([\\d.]+)\\s*\\)/\n              );\n              if (!alphaMatch || parseFloat(alphaMatch[1]) > 0.1) {\n                return c;\n              }\n            }\n          }\n        }\n      }\n      return "rgb(255, 255, 255)";\n    }\n    function extractTextStyle(style, includePadding = true) {\n      let textAlign = style.textAlign || "left";\n      if (textAlign === "start") textAlign = "left";\n      else if (textAlign === "end") textAlign = "right";\n      if (textAlign === "left" && style.justifyContent === "center")\n        textAlign = "center";\n      return {\n        color: style.color,\n        fontSize: parseFloat(style.fontSize) || 16,\n        fontFamily: style.fontFamily,\n        fontWeight: parseInt(style.fontWeight, 10) || 400,\n        textAlign,\n        lineHeight: parseFloat(style.lineHeight) || 0,\n        letterSpacing: parseFloat(style.letterSpacing) || 0,\n        ...includePadding && parseFloat(style.paddingTop) > 0 ? { paddingTop: parseFloat(style.paddingTop) } : {},\n        ...includePadding && parseFloat(style.paddingRight) > 0 ? { paddingRight: parseFloat(style.paddingRight) } : {},\n        ...includePadding && parseFloat(style.paddingBottom) > 0 ? { paddingBottom: parseFloat(style.paddingBottom) } : {},\n        ...includePadding && parseFloat(style.paddingLeft) > 0 ? { paddingLeft: parseFloat(style.paddingLeft) } : {},\n        ...style.whiteSpace === "nowrap" ? { whiteSpace: "nowrap" } : {}\n      };\n    }\n    function extractStackingMetadata(style) {\n      const positioned = style.position === "absolute" || style.position === "fixed" || style.position === "relative" || style.position === "sticky";\n      const parsed = Number.parseInt(style.zIndex, 10);\n      return {\n        ...positioned ? { stackingPositioned: true } : {},\n        ...positioned && Number.isFinite(parsed) ? { stackingIndex: parsed } : {}\n      };\n    }\n    function isSemanticInlineHighlightTag(tag) {\n      return tag === "strong" || tag === "mark" || tag === "code";\n    }\n    function extractEffectiveBg(elStyle) {\n      const bg = elStyle.backgroundColor;\n      const hasBg = !!bg && bg !== "transparent" && bg !== "rgba(0, 0, 0, 0)";\n      if (hasBg) {\n        const alphaMatch = bg.match(/,\\s*([\\d.]+)\\s*\\)$/);\n        if (!alphaMatch || parseFloat(alphaMatch[1]) !== 0) return bg;\n      }\n      const bi = elStyle.backgroundImage;\n      if (!bi || bi === "none" || !bi.includes("linear-gradient")) return void 0;\n      const colorMatches = bi.match(/rgba?\\([^)]+\\)/g);\n      if (!colorMatches || colorMatches.length === 0) return void 0;\n      for (let ci = colorMatches.length - 1; ci >= 0; ci--) {\n        const c = colorMatches[ci];\n        if (c !== "rgba(0, 0, 0, 0)" && c !== "rgba(0,0,0,0)" && !/rgba\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*0\\s*\\)/.test(c)) {\n          return c;\n        }\n      }\n      return void 0;\n    }\n    function isEmojiImg(imgEl) {\n      const alt = imgEl.alt ?? "";\n      return !!(imgEl.classList?.contains("emoji") || imgEl.src && (imgEl.src.includes("twemoji") || imgEl.src.includes("/emoji/")) || alt.length > 0 && alt.length <= 8 && /\\p{Extended_Pictographic}/u.test(alt));\n    }\n    function extractTextRuns(element, skipInlineBadges = false, stripBgBadges = false, skipBlockEls = false) {\n      const runs = [];\n      const elementStyle = getComputedStyle(element);\n      const elementBg = elementStyle.backgroundColor;\n      const elementHasBg = !(stripBgBadges instanceof Set && stripBgBadges.has(element)) && !!elementBg && elementBg !== "transparent" && elementBg !== "rgba(0, 0, 0, 0)";\n      function lastIsBreak() {\n        return runs.length > 0 && runs[runs.length - 1].breakLine === true;\n      }\n      function pushText(text, style, bg) {\n        const segments = text.split("\\n");\n        for (let i = 0; i < segments.length; i++) {\n          const seg = segments[i];\n          if (seg !== "") {\n            const run = {\n              text: seg,\n              color: style.color,\n              fontSize: parseFloat(style.fontSize) || 16,\n              fontFamily: style.fontFamily,\n              bold: parseInt(style.fontWeight, 10) >= 600,\n              italic: style.fontStyle === "italic",\n              underline: style.textDecorationLine?.includes("underline"),\n              strikethrough: style.textDecorationLine?.includes("line-through")\n            };\n            if (bg) run.backgroundColor = bg;\n            runs.push(run);\n          }\n          if (i < segments.length - 1 && !lastIsBreak()) {\n            runs.push({ text: "", breakLine: true });\n          }\n        }\n      }\n      for (const node of Array.from(element.childNodes)) {\n        if (node.nodeType === Node.TEXT_NODE) {\n          const text = node.textContent ?? "";\n          if (text.trim() === "") {\n            const newlineCount = (text.match(/\\n/g) ?? []).length;\n            for (let i = 0; i < newlineCount; i++) {\n              if (!lastIsBreak()) runs.push({ text: "", breakLine: true });\n            }\n            if (newlineCount === 0 && text.length > 0 && !lastIsBreak()) {\n              const ws = elementStyle.whiteSpace;\n              const preserve = ws === "pre" || ws === "pre-wrap" || ws === "break-spaces";\n              pushText(preserve ? text : " ", elementStyle, elementHasBg ? elementBg : void 0);\n            }\n            continue;\n          }\n          pushText(text, elementStyle, elementHasBg ? elementBg : void 0);\n        } else if (node.nodeType === Node.ELEMENT_NODE) {\n          const el = node;\n          const tag = el.tagName.toLowerCase();\n          if (tag === "br") {\n            if (!lastIsBreak()) runs.push({ text: "", breakLine: true });\n            continue;\n          }\n          if (el.classList?.contains("katex-mathml")) continue;\n          if (tag === "mjx-container" || tag === "svg") continue;\n          if (tag === "sub" || tag === "sup") {\n            const childRuns = extractTextRuns(el, skipInlineBadges, stripBgBadges);\n            const flag = tag === "sub" ? "subscript" : "superscript";\n            childRuns.forEach((r) => {\n              if (!r.breakLine) r[flag] = true;\n            });\n            runs.push(...childRuns);\n            continue;\n          }\n          if (tag === "a") {\n            const href = el.href;\n            const childRuns = extractTextRuns(el);\n            childRuns.forEach((r) => {\n              if (!r.breakLine) r.hyperlink = href;\n            });\n            runs.push(...childRuns);\n            continue;\n          }\n          if (tag === "img") {\n            const imgEl = el;\n            const alt = imgEl.alt ?? "";\n            if (isEmojiImg(imgEl) && alt) {\n              pushText(alt, getComputedStyle(el), void 0);\n            }\n            continue;\n          }\n          const elStyle = getComputedStyle(el);\n          if (/^(block|flex|grid|list-item|table)/.test(elStyle.display)) {\n            if (skipBlockEls instanceof Set && skipBlockEls.has(el)) continue;\n            if (!lastIsBreak() && runs.length > 0) {\n              runs.push({ text: "", breakLine: true });\n            }\n            runs.push(...extractTextRuns(el));\n          } else {\n            const bg = elStyle.backgroundColor;\n            const hasBg = bg && bg !== "transparent" && bg !== "rgba(0, 0, 0, 0)";\n            const alphaZero = hasBg && (() => {\n              const m = bg.match(/,\\s*([\\d.]+)\\s*\\)$/);\n              return m ? parseFloat(m[1]) === 0 : false;\n            })();\n            const effectiveBg = stripBgBadges instanceof Set && stripBgBadges.has(el) ? void 0 : extractEffectiveBg(elStyle);\n            const inlineElBorderRadius = parseFloat(elStyle.borderRadius) || 0;\n            const inlineTag = el.tagName.toLowerCase();\n            const isSemanticInlineHighlight = isSemanticInlineHighlightTag(inlineTag);\n            const isInlineEligibleBadge = (() => {\n              if (isSemanticInlineHighlight) return false;\n              if (elStyle.display !== "inline") return false;\n              if (inlineTag !== "span") return false;\n              if (inlineElBorderRadius <= 0) return false;\n              const m = (effectiveBg ?? "").match(/,\\s*([\\d.]+)\\s*\\)$/);\n              return m ? parseFloat(m[1]) >= 0.5 : true;\n            })();\n            const isInlineBadgeDisplay = !isSemanticInlineHighlight && (elStyle.display === "inline-block" || elStyle.display === "inline-flex" || elStyle.display === "inline-grid");\n            const isBadge = effectiveBg && !alphaZero && (isInlineBadgeDisplay || isInlineEligibleBadge);\n            if (isBadge) {\n              const shouldSkipBadge = skipInlineBadges === true || skipInlineBadges instanceof Set && skipInlineBadges.has(el);\n              if (shouldSkipBadge) {\n                continue;\n              }\n              const shouldStripBg = stripBgBadges instanceof Set && stripBgBadges.has(el);\n              const childRuns2 = extractTextRuns(el, false);\n              if (shouldStripBg) {\n                childRuns2.forEach((r) => {\n                  if (!r.breakLine) r.backgroundColor = void 0;\n                });\n              } else {\n                childRuns2.forEach((r) => {\n                  if (!r.breakLine && !r.backgroundColor) r.backgroundColor = effectiveBg;\n                });\n              }\n              runs.push(...childRuns2);\n              continue;\n            }\n            const childRuns = extractTextRuns(el, skipInlineBadges, stripBgBadges);\n            if (effectiveBg) {\n              childRuns.forEach((r) => {\n                if (!r.breakLine && !r.backgroundColor) r.backgroundColor = effectiveBg;\n              });\n            }\n            runs.push(...childRuns);\n          }\n        }\n      }\n      while (runs.length > 0 && runs[runs.length - 1].breakLine) {\n        runs.pop();\n      }\n      while (runs.length > 0 && runs[0].breakLine) {\n        runs.shift();\n      }\n      return runs;\n    }\n    function computeLeadingOffset(badgeShapes, containerRect, slideRect) {\n      if (badgeShapes.length === 0) return 0;\n      const containerSSLeft = containerRect.left - slideRect.left;\n      const leading = badgeShapes.filter((b) => {\n        if (b.x > containerSSLeft + 8) return false;\n        const runs = b.runs;\n        return Array.isArray(runs) && runs.length > 0;\n      });\n      if (leading.length === 0) return 0;\n      const rightEdge = leading.reduce(\n        (max, b) => Math.max(max, b.x + b.width),\n        containerSSLeft\n      );\n      return Math.max(0, rightEdge - containerSSLeft);\n    }\n    function computeShallowFlexOffset(container, containerRect, slideRect, style) {\n      const containerSSLeft = containerRect.left - slideRect.left;\n      const containerGap = parseFloat(style.columnGap || style.gap || "0") || 0;\n      for (const node of Array.from(container.childNodes)) {\n        if (node.nodeType === Node.TEXT_NODE) {\n          if ((node.textContent ?? "").trim() !== "") return 0;\n          continue;\n        }\n        if (node.nodeType !== Node.ELEMENT_NODE) continue;\n        const nodeEl = node;\n        const nodeStyle = getComputedStyle(nodeEl);\n        if (nodeStyle.display === "none" || nodeStyle.visibility === "hidden") {\n          continue;\n        }\n        const nodeRect = nodeEl.getBoundingClientRect();\n        if (nodeRect.width <= 0 || nodeRect.height <= 0) continue;\n        return Math.max(0, nodeRect.right - slideRect.left - containerSSLeft) + containerGap;\n      }\n      return 0;\n    }\n    function extractListItemEl(li, level = 0, skipBadges = false, stripBadges = false, leadingOffsetPx = 0) {\n      const items = [];\n      const runs = [];\n      const nestedItems = [];\n      function lastIsBreak() {\n        return runs.length > 0 && runs[runs.length - 1].breakLine === true;\n      }\n      for (const node of Array.from(li.childNodes)) {\n        if (node.nodeType === Node.TEXT_NODE) {\n          const text = node.textContent ?? "";\n          if (text.trim() === "") continue;\n          const liStyle = getComputedStyle(li);\n          const liBg = liStyle.backgroundColor;\n          const liHasBg = !!liBg && liBg !== "transparent" && liBg !== "rgba(0, 0, 0, 0)";\n          const segments = text.split("\\n");\n          for (let i = 0; i < segments.length; i++) {\n            const seg = segments[i];\n            if (seg !== "") {\n              const run = {\n                text: seg,\n                color: liStyle.color,\n                fontSize: parseFloat(liStyle.fontSize) || 16,\n                fontFamily: liStyle.fontFamily,\n                bold: parseInt(liStyle.fontWeight, 10) >= 600,\n                italic: liStyle.fontStyle === "italic"\n              };\n              if (liHasBg) run.backgroundColor = liBg;\n              runs.push(run);\n            }\n            if (i < segments.length - 1) {\n              if (!lastIsBreak()) runs.push({ text: "", breakLine: true });\n            }\n          }\n        } else if (node.nodeType === Node.ELEMENT_NODE) {\n          const el = node;\n          const childTag = el.tagName.toLowerCase();\n          if (childTag === "ul" || childTag === "ol") {\n            nestedItems.push(...extractListItems(el, level + 1));\n          } else if (childTag === "br") {\n            if (!lastIsBreak()) runs.push({ text: "", breakLine: true });\n          } else if (childTag === "img") {\n            const imgEl = el;\n            const alt = imgEl.alt ?? "";\n            if (isEmojiImg(imgEl) && alt) {\n              const liStyle = getComputedStyle(li);\n              runs.push({\n                text: alt,\n                color: liStyle.color,\n                fontSize: parseFloat(liStyle.fontSize) || 16,\n                fontFamily: liStyle.fontFamily,\n                bold: parseInt(liStyle.fontWeight, 10) >= 600,\n                italic: liStyle.fontStyle === "italic"\n              });\n            }\n          } else if (childTag === "sub" || childTag === "sup") {\n            const childRuns = extractTextRuns(el, skipBadges, stripBadges);\n            const flag = childTag === "sub" ? "subscript" : "superscript";\n            childRuns.forEach((r) => {\n              if (!r.breakLine) r[flag] = true;\n            });\n            runs.push(...childRuns);\n          } else {\n            if (skipBadges !== false && skipBadges.has(el)) continue;\n            const isBlockChild = /^(p|div|blockquote|pre|figure|h[1-6]|section|article|aside|header|footer|main)$/.test(childTag);\n            if (isBlockChild && runs.length > 0 && !lastIsBreak()) {\n              runs.push({ text: "", breakLine: true });\n            }\n            if (childTag === "pre" || childTag === "marp-pre") {\n              continue;\n            }\n            const childRuns = extractTextRuns(el, skipBadges, stripBadges);\n            if (stripBadges instanceof Set && stripBadges.has(el)) {\n              childRuns.forEach((r) => {\n                if (!r.breakLine) r.backgroundColor = void 0;\n              });\n            } else if (!isBlockChild) {\n              const elBg = extractEffectiveBg(getComputedStyle(el));\n              if (elBg) {\n                childRuns.forEach((r) => {\n                  if (!r.breakLine && !r.backgroundColor) r.backgroundColor = elBg;\n                });\n              }\n            }\n            runs.push(...childRuns);\n          }\n        }\n      }\n      if (runs.length > 0) {\n        const combinedText = runs.map((r) => r.text).join("");\n        const liStyle = getComputedStyle(li);\n        items.push({\n          text: combinedText.trim(),\n          level,\n          listStyleType: liStyle.listStyleType || void 0,\n          runs,\n          ...leadingOffsetPx > 0 ? { leadingOffset: leadingOffsetPx } : {}\n        });\n      }\n      items.push(...nestedItems);\n      return items;\n    }\n    function extractListItems(list, level = 0, perLiSkipMap, perLiStripMap, perLiLeadingOffsetMap) {\n      const items = [];\n      for (const child of Array.from(list.children)) {\n        if (child.tagName.toLowerCase() === "li") {\n          const skipBadges = perLiSkipMap?.get(child) ?? false;\n          const stripBadges = perLiStripMap?.get(child) ?? false;\n          const leadingOffsetPx = perLiLeadingOffsetMap?.get(child) ?? 0;\n          items.push(\n            ...extractListItemEl(\n              child,\n              level,\n              skipBadges,\n              stripBadges,\n              leadingOffsetPx\n            )\n          );\n        }\n      }\n      return items;\n    }\n    function extractCodeRuns(codeEl) {\n      const runs = [];\n      const defaultStyle = getComputedStyle(codeEl);\n      function lastIsBreak() {\n        return runs.length > 0 && runs[runs.length - 1].breakLine === true;\n      }\n      function walk(node) {\n        if (node.nodeType === Node.TEXT_NODE) {\n          const text = node.textContent ?? "";\n          if (text === "") return;\n          const parent = node.parentElement ?? codeEl;\n          const style = getComputedStyle(parent);\n          const segments = text.split("\\n");\n          for (let i = 0; i < segments.length; i++) {\n            const seg = segments[i];\n            if (seg !== "") {\n              runs.push({\n                text: seg,\n                color: style.color,\n                fontSize: parseFloat(style.fontSize) || parseFloat(defaultStyle.fontSize) || 16,\n                fontFamily: style.fontFamily || defaultStyle.fontFamily,\n                bold: parseInt(style.fontWeight, 10) >= 600,\n                italic: style.fontStyle === "italic"\n              });\n            }\n            if (i < segments.length - 1) {\n              runs.push({ text: "", breakLine: true });\n            }\n          }\n        } else if (node.nodeType === Node.ELEMENT_NODE) {\n          for (const child of Array.from(node.childNodes)) {\n            walk(child);\n          }\n        }\n      }\n      walk(codeEl);\n      return runs;\n    }\n    function extractTableData(table) {\n      const rows = [];\n      const colWidths = [];\n      for (const tr of Array.from(table.querySelectorAll("tr"))) {\n        const cells = [];\n        const isFirstRow = rows.length === 0;\n        const trStyle = getComputedStyle(tr);\n        const trBg = trStyle.backgroundColor;\n        const trHasBg = !!trBg && trBg !== "transparent" && trBg !== "rgba(0, 0, 0, 0)";\n        for (const td of Array.from(tr.querySelectorAll("th, td"))) {\n          const style = getComputedStyle(td);\n          const tdColSpan = td.colSpan || 1;\n          const tdRowSpan = td.rowSpan || 1;\n          if (isFirstRow) {\n            const perColWidth = td.offsetWidth / tdColSpan;\n            for (let i = 0; i < tdColSpan; i++) {\n              colWidths.push(perColWidth);\n            }\n          }\n          const cellBg = style.backgroundColor;\n          const cellHasBg = !!cellBg && cellBg !== "transparent" && cellBg !== "rgba(0, 0, 0, 0)";\n          const effectiveBg = cellHasBg ? cellBg : trHasBg ? trBg : cellBg;\n          const cellRuns = extractTextRuns(td);\n          const paragraphs = [{ runs: [] }];\n          for (const run of cellRuns) {\n            if (run.breakLine) {\n              paragraphs.push({ runs: [] });\n            } else {\n              paragraphs[paragraphs.length - 1].runs.push(run);\n            }\n          }\n          if (paragraphs.length > 1 && paragraphs[paragraphs.length - 1].runs.length === 0) {\n            paragraphs.pop();\n          }\n          cells.push({\n            text: td.textContent ?? "",\n            runs: cellRuns,\n            paragraphs,\n            isHeader: td.tagName.toLowerCase() === "th",\n            ...tdColSpan > 1 ? { colspan: tdColSpan } : {},\n            ...tdRowSpan > 1 ? { rowspan: tdRowSpan } : {},\n            style: {\n              color: style.color,\n              backgroundColor: effectiveBg,\n              fontSize: parseFloat(style.fontSize) || 16,\n              fontFamily: style.fontFamily,\n              fontWeight: parseInt(style.fontWeight, 10) || 400,\n              textAlign: style.textAlign || "left",\n              borderColor: style.borderColor,\n              borderWidth: parseFloat(style.borderTopWidth) || 0,\n              ...style.borderTopStyle && style.borderTopStyle !== "none" ? { borderStyle: style.borderTopStyle } : {},\n              paddingTop: parseFloat(style.paddingTop) || 0,\n              paddingRight: parseFloat(style.paddingRight) || 0,\n              paddingBottom: parseFloat(style.paddingBottom) || 0,\n              paddingLeft: parseFloat(style.paddingLeft) || 0\n            }\n          });\n        }\n        rows.push({\n          cells,\n          ...tr.getBoundingClientRect().height > 0 ? { height: tr.getBoundingClientRect().height } : {}\n        });\n      }\n      return { rows, colWidths };\n    }\n    function extractBadgeTextLines(element, slideRect) {\n      const lines = [];\n      const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);\n      let node;\n      let measuredAnyCharacter = false;\n      while (node = walker.nextNode()) {\n        const parent = node.parentElement;\n        const text = node.textContent ?? "";\n        if (!parent || !text) continue;\n        const style = getComputedStyle(parent);\n        if (style.display === "none" || style.visibility === "hidden") continue;\n        let offset = 0;\n        for (const character of Array.from(text)) {\n          const characterLength = character.length;\n          const range = document.createRange();\n          if (typeof range.setStart !== "function" || typeof range.setEnd !== "function" || typeof range.getBoundingClientRect !== "function") {\n            return void 0;\n          }\n          range.setStart(node, offset);\n          range.setEnd(node, offset + characterLength);\n          offset += characterLength;\n          const rect = range.getBoundingClientRect();\n          if (rect.width <= 0 || rect.height <= 0) continue;\n          measuredAnyCharacter = true;\n          const topTolerance = Math.max(\n            0.5,\n            (parseFloat(style.lineHeight) || rect.height) * 0.2\n          );\n          let line = lines.find(\n            (candidate) => Math.abs(candidate.y - (rect.top - slideRect.top)) <= topTolerance\n          );\n          if (!line) {\n            line = {\n              x: rect.left - slideRect.left,\n              y: rect.top - slideRect.top,\n              width: rect.width,\n              height: rect.height,\n              runs: [],\n              lineHeight: parseFloat(style.lineHeight) || void 0,\n              textAlign: style.textAlign || "left"\n            };\n            lines.push(line);\n          } else {\n            const right = line.x + line.width;\n            const bottom = line.y + line.height;\n            line.x = Math.min(line.x, rect.left - slideRect.left);\n            line.y = Math.min(line.y, rect.top - slideRect.top);\n            line.width = Math.max(right, rect.right - slideRect.left) - line.x;\n            line.height = Math.max(bottom, rect.bottom - slideRect.top) - line.y;\n          }\n          line.runs.push({\n            text: character,\n            color: style.color,\n            fontSize: parseFloat(style.fontSize) || 16,\n            fontFamily: style.fontFamily,\n            bold: parseInt(style.fontWeight, 10) >= 600,\n            italic: style.fontStyle === "italic",\n            underline: style.textDecorationLine?.includes("underline"),\n            strikethrough: style.textDecorationLine?.includes("line-through")\n          });\n        }\n      }\n      for (const image of Array.from(\n        element.querySelectorAll("img")\n      )) {\n        if (!isEmojiImg(image)) continue;\n        const rect = image.getBoundingClientRect();\n        if (rect.width <= 0 || rect.height <= 0) continue;\n        const imageTop = rect.top - slideRect.top;\n        const line = lines.find(\n          (candidate) => Math.abs(candidate.y - imageTop) <= Math.max(candidate.lineHeight ?? rect.height, rect.height) * 0.5\n        );\n        if (!line || !image.alt) continue;\n        line.width = Math.max(line.x + line.width, rect.right - slideRect.left) - line.x;\n        line.height = Math.max(line.y + line.height, rect.bottom - slideRect.top) - line.y;\n        line.runs.push({\n          text: image.alt,\n          color: getComputedStyle(image).color,\n          fontSize: parseFloat(getComputedStyle(image).fontSize) || 16,\n          fontFamily: getComputedStyle(image).fontFamily\n        });\n        measuredAnyCharacter = true;\n      }\n      return measuredAnyCharacter ? lines.sort((left, right) => left.y - right.y) : void 0;\n    }\n    function extractMeasuredParagraphLayout(element, fallbackRuns, slideRect) {\n      if (element.querySelector("img,svg,[data-bespoke-marp-fragment]"))\n        return void 0;\n      const unsupportedBlock = Array.from(element.children).some(\n        (child) => /^(block|flex|grid|list-item|table)/.test(\n          getComputedStyle(child).display\n        )\n      );\n      if (unsupportedBlock) return void 0;\n      const fallbackText = fallbackRuns.filter((run) => !run.breakLine).map((run) => run.text).join("");\n      if (!/[\\u3000-\\u30ff\\u3400-\\u9fff\\uf900-\\ufaff\\uff01-\\uff9f]/.test(\n        fallbackText\n      ))\n        return void 0;\n      const lines = [];\n      const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);\n      let node;\n      while (node = walker.nextNode()) {\n        const text = node.textContent ?? "";\n        const parent = node.parentElement;\n        if (!parent || !text) continue;\n        const style = getComputedStyle(parent);\n        let offset = 0;\n        for (const character of Array.from(text)) {\n          const charLength = character.length;\n          const range = document.createRange();\n          if (typeof range.setStart !== "function" || typeof range.setEnd !== "function" || typeof range.getBoundingClientRect !== "function")\n            return void 0;\n          range.setStart(node, offset);\n          range.setEnd(node, offset + charLength);\n          offset += charLength;\n          const rect = range.getBoundingClientRect();\n          if (rect.width <= 0 || rect.height <= 0) continue;\n          let line = lines[lines.length - 1];\n          if (!line || Math.abs(rect.top - line.y) > Math.max(1, rect.height * 0.25)) {\n            line = {\n              x: rect.left,\n              y: rect.top,\n              width: rect.width,\n              height: rect.height,\n              runs: [],\n              lineHeight: parseFloat(style.lineHeight) || rect.height\n            };\n            lines.push(line);\n          } else {\n            line.width = Math.max(line.x + line.width, rect.right) - line.x;\n            line.height = Math.max(line.height, rect.bottom - line.y);\n          }\n          const last = line.runs[line.runs.length - 1];\n          const styleKey = `${style.color}|${style.fontSize}|${style.fontFamily}|${style.fontWeight}|${style.fontStyle}|${parent.closest("a")?.href ?? ""}|${extractEffectiveBg(style) ?? ""}`;\n          if (last && last.__styleKey === styleKey) {\n            last.text += character;\n          } else {\n            line.runs.push({\n              text: character,\n              color: style.color,\n              fontSize: parseFloat(style.fontSize) || 16,\n              fontFamily: style.fontFamily,\n              bold: parseInt(style.fontWeight, 10) >= 600,\n              italic: style.fontStyle === "italic",\n              underline: style.textDecorationLine?.includes("underline"),\n              strikethrough: style.textDecorationLine?.includes("line-through"),\n              hyperlink: parent.closest("a")?.href || void 0,\n              backgroundColor: extractEffectiveBg(style),\n              subscript: parent.tagName.toLowerCase() === "sub" || parent.closest("sub") !== null || void 0,\n              superscript: parent.tagName.toLowerCase() === "sup" || parent.closest("sup") !== null || void 0,\n              __styleKey: styleKey\n            });\n          }\n        }\n      }\n      const measuredText = lines.flatMap((line) => line.runs).map((run) => run.text).join("");\n      if (measuredText !== fallbackText) return void 0;\n      const cleanLines = lines.map((line) => {\n        const runs2 = line.runs.map((run) => {\n          delete run.__styleKey;\n          return run;\n        });\n        const align = fallbackRuns.length > 0 ? getComputedStyle(element).textAlign : void 0;\n        return {\n          x: line.x - slideRect.left,\n          y: line.y - slideRect.top,\n          width: line.width,\n          height: line.height,\n          runs: runs2,\n          lineHeight: line.lineHeight,\n          ...align === "center" || align === "right" || align === "justify" ? {\n            textAlign: align\n          } : {}\n        };\n      });\n      const runs = cleanLines.flatMap((line, index) => [\n        ...line.runs,\n        ...index < cleanLines.length - 1 ? [{ text: "", breakLine: true }] : []\n      ]);\n      return { runs, lines: cleanLines };\n    }\n    function extractMeasuredParagraphRuns(element, fallbackRuns, slideRect) {\n      return extractMeasuredParagraphLayout(element, fallbackRuns, slideRect)?.runs;\n    }\n    function extractInlineDecorationShapes(container, slideRect) {\n      const shapes = [];\n      const elements = [];\n      for (const element of Array.from(\n        container.querySelectorAll("strong,mark,code,span")\n      )) {\n        const style = getComputedStyle(element);\n        const tag = element.tagName.toLowerCase();\n        const hasGradient = style.backgroundImage.includes("linear-gradient");\n        if ((tag === "code" || tag === "span") && !hasGradient)\n          continue;\n        const backgroundColor = extractEffectiveBg(style);\n        if (!backgroundColor) continue;\n        if (tag === "span" && style.display !== "inline")\n          continue;\n        const range = document.createRange();\n        range.selectNodeContents(element);\n        if (typeof range.getBoundingClientRect !== "function") continue;\n        const rect = range.getBoundingClientRect();\n        if (rect.width <= 0 || rect.height <= 0) continue;\n        const gradient = style.backgroundImage;\n        const markerStops = Array.from(\n          gradient.matchAll(/(\\d+(?:\\.\\d+)?)%/g)\n        ).map((match) => Number(match[1]));\n        const markerStop = gradient.includes("linear-gradient") && markerStops.length > 0 ? Math.min(90, ...markerStops) : 0;\n        const y = rect.top - slideRect.top + rect.height * markerStop / 100;\n        const borderRadius = parseFloat(style.borderRadius) || 0;\n        shapes.push({\n          type: "container",\n          children: [],\n          x: rect.left - slideRect.left,\n          y,\n          width: rect.width,\n          height: Math.max(1, rect.height * (1 - markerStop / 100)),\n          style: {\n            backgroundColor,\n            ...borderRadius > 0 ? { borderRadius } : {}\n          },\n          source: sourceRefFor(element)\n        });\n        elements.push(element);\n      }\n      return { shapes, elements };\n    }\n    function extractInlineBadgeShapes(container, slideRect, containerRect) {\n      const badges = [];\n      const badgeEls = [];\n      const bgOnlyElements = [];\n      const containerSSLeft = containerRect ? containerRect.left - slideRect.left : -Infinity;\n      const containerHasNonBadgeText = (() => {\n        for (const node of Array.from(container.childNodes)) {\n          if (node.nodeType === Node.TEXT_NODE) {\n            if ((node.textContent ?? "").trim() !== "") return true;\n            continue;\n          }\n          if (node.nodeType !== Node.ELEMENT_NODE) continue;\n          const cs = getComputedStyle(node);\n          if (cs.display === "none" || cs.visibility === "hidden") continue;\n          const isChildBadge = cs.display === "inline-block" || cs.display === "inline-flex" || cs.display === "inline-grid";\n          if (!isChildBadge) return true;\n        }\n        return false;\n      })();\n      for (const el of Array.from(container.querySelectorAll("*"))) {\n        const s = getComputedStyle(el);\n        const inlineTag = el.tagName.toLowerCase();\n        const isSemanticInlineHighlight = isSemanticInlineHighlightTag(inlineTag);\n        const isInlineBadgeDisplay = !isSemanticInlineHighlight && (s.display === "inline-block" || s.display === "inline-flex" || s.display === "inline-grid");\n        const inlineBorderRadius = parseFloat(s.borderRadius) || 0;\n        const isInlineWithRoundedBg = !isSemanticInlineHighlight && s.display === "inline" && inlineTag === "span" && inlineBorderRadius > 0;\n        if (!isInlineBadgeDisplay && !isInlineWithRoundedBg) continue;\n        const bg = s.backgroundColor;\n        if (!bg || bg === "transparent") continue;\n        const alphaMatch = bg.match(/,\\s*([\\d.]+)\\s*\\)$/);\n        if (alphaMatch && parseFloat(alphaMatch[1]) === 0) continue;\n        if (isInlineWithRoundedBg && alphaMatch && parseFloat(alphaMatch[1]) < 0.5) continue;\n        const iRect = el.getBoundingClientRect();\n        if (iRect.width === 0 || iRect.height === 0) continue;\n        if (isInlineBadgeDisplay && containerRect && containerHasNonBadgeText && iRect.left - slideRect.left > containerSSLeft + 8) {\n        }\n        const br = inlineBorderRadius;\n        const badgeRuns = extractTextRuns(el);\n        const textLines = extractBadgeTextLines(el, slideRect);\n        let textBounds;\n        const textRange = document.createRange();\n        textRange.selectNodeContents(el);\n        if (typeof textRange.getBoundingClientRect === "function") {\n          const textRect = textRange.getBoundingClientRect();\n          if (textRect.left < iRect.left - 0.5 || textRect.right > iRect.right + 0.5 || textRect.top < iRect.top - 0.5 || textRect.bottom > iRect.bottom + 0.5) {\n            textBounds = {\n              x: textRect.left - slideRect.left,\n              y: textRect.top - slideRect.top,\n              width: Math.max(1, textRect.width),\n              height: Math.max(1, textRect.height),\n              lineHeight: parseFloat(s.lineHeight) || void 0\n            };\n          }\n        }\n        badgeRuns.forEach((r) => {\n          if (!r.breakLine) r.backgroundColor = void 0;\n        });\n        const hasBadgeText = badgeRuns.some(\n          (r) => !r.breakLine && r.text.trim() !== ""\n        );\n        badges.push({\n          type: "container",\n          children: [],\n          ...hasBadgeText ? { runs: badgeRuns } : {},\n          ...textLines?.length ? { textLines } : {},\n          ...textBounds ? { textBounds } : {},\n          x: iRect.left - slideRect.left,\n          y: iRect.top - slideRect.top,\n          width: iRect.width,\n          height: iRect.height,\n          style: {\n            backgroundColor: bg,\n            ...br > 0 ? { borderRadius: br } : {}\n          }\n        });\n        badgeEls.push(el);\n      }\n      return { shapes: badges, elements: badgeEls, bgOnlyElements };\n    }\n    function extractInlineTextSegments(container, slideRect, badgeElements, badgeShapesByElement) {\n      const segments = [];\n      function walk(node) {\n        if (node.nodeType === Node.ELEMENT_NODE) {\n          const element = node;\n          if (element.tagName.toLowerCase() === "svg" || element.tagName.toLowerCase() === "mjx-container") {\n            return;\n          }\n          const badgeShape = badgeShapesByElement?.get(element);\n          if (badgeShape) {\n            segments.push(badgeShape);\n            return;\n          }\n          if (badgeElements.has(element)) return;\n          for (const child of Array.from(element.childNodes)) walk(child);\n          return;\n        }\n        if (node.nodeType !== Node.TEXT_NODE) return;\n        const parent = node.parentElement;\n        const text = node.textContent ?? "";\n        if (!parent || text.trim() === "") return;\n        const range = document.createRange();\n        range.selectNodeContents(node);\n        if (typeof range.getBoundingClientRect !== "function") return;\n        const rect = range.getBoundingClientRect();\n        if (rect.width <= 0 || rect.height <= 0) return;\n        const style = getComputedStyle(parent);\n        const source = sourceRefFor(parent);\n        segments.push({\n          type: "paragraph",\n          source: {\n            ...source,\n            id: `${source.id}:text`\n          },\n          x: rect.left - slideRect.left,\n          y: rect.top - slideRect.top,\n          width: rect.width,\n          height: rect.height,\n          runs: [\n            {\n              text: text.trim(),\n              color: style.color,\n              fontSize: parseFloat(style.fontSize) || 16,\n              fontFamily: style.fontFamily,\n              bold: parseInt(style.fontWeight, 10) >= 600,\n              italic: style.fontStyle === "italic",\n              underline: style.textDecorationLine?.includes("underline"),\n              strikethrough: style.textDecorationLine?.includes("line-through")\n            }\n          ],\n          style: extractTextStyle(style, false)\n        });\n      }\n      for (const child of Array.from(container.childNodes)) walk(child);\n      return segments;\n    }\n    function resolveClippedRect(element, sourceRect, slideRect) {\n      let left = sourceRect.left;\n      let top = sourceRect.top;\n      let right = sourceRect.right;\n      let bottom = sourceRect.bottom;\n      let ancestor = element.parentElement;\n      let isClipped = false;\n      while (ancestor && ancestor.getBoundingClientRect) {\n        const ancestorStyle = getComputedStyle(ancestor);\n        const clips = ancestorStyle.overflow === "hidden" || ancestorStyle.overflow === "clip" || ancestorStyle.overflowX === "hidden" || ancestorStyle.overflowX === "clip" || ancestorStyle.overflowY === "hidden" || ancestorStyle.overflowY === "clip";\n        if (clips) {\n          const ancestorRect = ancestor.getBoundingClientRect();\n          const nextLeft = Math.max(left, ancestorRect.left);\n          const nextTop = Math.max(top, ancestorRect.top);\n          const nextRight = Math.min(right, ancestorRect.right);\n          const nextBottom = Math.min(bottom, ancestorRect.bottom);\n          isClipped ||= nextLeft !== left || nextTop !== top || nextRight !== right || nextBottom !== bottom;\n          left = nextLeft;\n          top = nextTop;\n          right = nextRight;\n          bottom = nextBottom;\n        }\n        if (ancestor === document.documentElement) break;\n        ancestor = ancestor.parentElement;\n      }\n      const slideLeft = Math.max(left, slideRect.left);\n      const slideTop = Math.max(top, slideRect.top);\n      const slideRight = Math.min(right, slideRect.right);\n      const slideBottom = Math.min(bottom, slideRect.bottom);\n      isClipped ||= slideLeft !== left || slideTop !== top || slideRight !== right || slideBottom !== bottom;\n      return {\n        left: slideLeft,\n        top: slideTop,\n        width: Math.max(0, slideRight - slideLeft),\n        height: Math.max(0, slideBottom - slideTop),\n        isClipped\n      };\n    }\n    function extractNestedImages(el, slideRect) {\n      const images = [];\n      for (const img of Array.from(el.querySelectorAll("img"))) {\n        const imgEl = img;\n        const rect = imgEl.getBoundingClientRect();\n        if (rect.width === 0 || rect.height === 0) continue;\n        const s = getComputedStyle(imgEl);\n        if (s.display === "none" || s.visibility === "hidden") continue;\n        if (isEmojiImg(imgEl)) continue;\n        const cssFilter = s.filter && s.filter !== "none" ? s.filter : void 0;\n        const clipped = resolveClippedRect(imgEl, rect, slideRect);\n        if (clipped.width <= 0 || clipped.height <= 0) continue;\n        images.push({\n          type: "image",\n          src: imgEl.src,\n          naturalWidth: imgEl.naturalWidth,\n          naturalHeight: imgEl.naturalHeight,\n          x: clipped.left - slideRect.left,\n          y: clipped.top - slideRect.top,\n          width: clipped.width,\n          height: clipped.height,\n          ...clipped.isClipped ? { rasterize: true } : cssFilter ? { cssFilter, pageX: rect.left, pageY: rect.top } : {}\n        });\n      }\n      for (const svg of Array.from(el.querySelectorAll("svg"))) {\n        const rect = svg.getBoundingClientRect();\n        if (rect.width === 0 || rect.height === 0) continue;\n        const clipped = resolveClippedRect(svg, rect, slideRect);\n        if (clipped.width <= 0 || clipped.height <= 0) continue;\n        try {\n          const svgStr = new XMLSerializer().serializeToString(svg);\n          const b64 = btoa(unescape(encodeURIComponent(svgStr)));\n          const dataUrl = `data:image/svg+xml;base64,${b64}`;\n          images.push({\n            type: "image",\n            src: dataUrl,\n            naturalWidth: clipped.width,\n            naturalHeight: clipped.height,\n            x: clipped.left - slideRect.left,\n            y: clipped.top - slideRect.top,\n            width: clipped.width,\n            height: clipped.height,\n            rasterize: true\n          });\n        } catch {\n        }\n      }\n      return images;\n    }\n    function extractCssBorders(style) {\n      const side = (width, borderStyle, color) => {\n        const parsedWidth = parseFloat(width) || 0;\n        if (parsedWidth <= 0 || borderStyle === "none") return void 0;\n        return {\n          width: parsedWidth,\n          color,\n          ...borderStyle && borderStyle !== "solid" ? { style: borderStyle } : {}\n        };\n      };\n      const top = side(style.borderTopWidth, style.borderTopStyle, style.borderTopColor);\n      const right = side(style.borderRightWidth, style.borderRightStyle, style.borderRightColor);\n      const bottom = side(style.borderBottomWidth, style.borderBottomStyle, style.borderBottomColor);\n      const left = side(style.borderLeftWidth, style.borderLeftStyle, style.borderLeftColor);\n      return {\n        ...top ? { top } : {},\n        ...right ? { right } : {},\n        ...bottom ? { bottom } : {},\n        ...left ? { left } : {}\n      };\n    }\n    function extractAvailableFlowWidth(element, rect) {\n      const parent = element.parentElement;\n      if (!parent) return rect.width;\n      const parentRect = parent.getBoundingClientRect();\n      if (parentRect.width <= 0 || parentRect.right < rect.left) return rect.width;\n      const rightEdge = Array.from(parent.children).filter((sibling) => sibling !== element).map((sibling) => sibling.getBoundingClientRect()).filter(\n        (siblingRect) => siblingRect.width > 0 && siblingRect.height > 0 && siblingRect.left >= rect.right - 1 && siblingRect.top < rect.bottom && siblingRect.bottom > rect.top\n      ).reduce(\n        (right, siblingRect) => Math.min(right, siblingRect.left),\n        parentRect.right\n      );\n      return Math.max(1, rightEdge - rect.left);\n    }\n    function walkElement(parent, child, slideRect) {\n      const elements = [];\n      const style = getComputedStyle(child);\n      if (style.display === "none") return [];\n      if (child.dataset?.marpitPresenterNotes !== void 0)\n        return [];\n      if (child.dataset?.marpitAdvancedBackgroundContainer !== void 0)\n        return [];\n      if (style.display === "contents") {\n        return walkDisplayContents(child, parent, slideRect);\n      }\n      if (style.visibility === "hidden") return [];\n      const rect = child.getBoundingClientRect();\n      if (rect.width === 0 || rect.height === 0) return [];\n      const tag = child.tagName.toLowerCase();\n      const parentIsFlexOrGrid = /^(flex|inline-flex|grid|inline-grid)/.test(\n        getComputedStyle(parent).display\n      );\n      if (!parentIsFlexOrGrid && tag !== "img" && tag !== "svg" && style.display === "inline")\n        return [];\n      const base = {\n        source: sourceRefFor(child),\n        x: rect.left - slideRect.left,\n        y: rect.top - slideRect.top,\n        width: rect.width,\n        height: rect.height,\n        ...extractStackingMetadata(style)\n      };\n      if (/^h[1-6]$/.test(tag)) {\n        const borderBottomWidth = parseFloat(style.borderBottomWidth) || 0;\n        const borderBottomStyle = style.borderBottomStyle;\n        const borderLeftWidth = parseFloat(style.borderLeftWidth) || 0;\n        const { shapes: headingBadgeShapes, elements: headingBadgeEls, bgOnlyElements: headingBgOnlyEls } = extractInlineBadgeShapes(child, slideRect, rect);\n        const headingLeadingOffset = computeLeadingOffset(\n          headingBadgeShapes,\n          rect,\n          slideRect\n        );\n        if (headingBadgeShapes.length > 0) elements.push(...headingBadgeShapes);\n        const headingBadgeSet = headingBadgeEls.length > 0 ? new Set(headingBadgeEls) : false;\n        const headingBgOnlySet = headingBgOnlyEls.length > 0 ? new Set(headingBgOnlyEls) : false;\n        const headingPaddingTop = parseFloat(style.paddingTop) || 0;\n        const headingPaddingRight = parseFloat(style.paddingRight) || 0;\n        const headingPaddingBottom = parseFloat(style.paddingBottom) || 0;\n        const headingPaddingLeft = parseFloat(style.paddingLeft) || 0;\n        const headingRuns = extractTextRuns(child, headingBadgeSet, headingBgOnlySet);\n        if (headingBadgeShapes.length === 0 || headingRuns.some((r) => !r.breakLine && r.text.trim() !== "")) {\n          elements.push({\n            type: "heading",\n            level: parseInt(tag[1], 10),\n            runs: headingRuns,\n            ...base,\n            x: base.x + headingLeadingOffset,\n            width: Math.max(10, base.width - headingLeadingOffset),\n            style: {\n              ...extractTextStyle(style),\n              ...headingPaddingTop || headingPaddingRight || headingPaddingBottom || headingPaddingLeft ? { paddingTop: headingPaddingTop, paddingRight: headingPaddingRight, paddingBottom: headingPaddingBottom, paddingLeft: headingPaddingLeft } : {}\n            },\n            ...borderBottomWidth > 0 ? {\n              borderBottom: {\n                width: borderBottomWidth,\n                color: style.borderBottomColor,\n                ...borderBottomStyle && borderBottomStyle !== "solid" && borderBottomStyle !== "none" ? { style: borderBottomStyle } : {}\n              }\n            } : {},\n            ...borderLeftWidth > 0 ? {\n              borderLeft: {\n                width: borderLeftWidth,\n                color: style.borderLeftColor\n              }\n            } : {}\n          });\n        }\n        elements.push(...extractNestedImages(child, slideRect));\n      } else if (tag === "p") {\n        if (child.querySelector("mjx-container")) {\n          const pRect = child.getBoundingClientRect();\n          if (pRect.width > 0 && pRect.height > 0) {\n            elements.push({\n              type: "image",\n              src: "",\n              // placeholder \u2014 rasterizeSlideTargets fills this via screenshot\n              naturalWidth: pRect.width,\n              naturalHeight: pRect.height,\n              x: pRect.left - slideRect.left,\n              y: pRect.top - slideRect.top,\n              width: pRect.width,\n              height: pRect.height,\n              rasterize: true\n            });\n          }\n          return elements;\n        }\n        const { shapes: paraBadgeShapes, elements: paraBadgeEls, bgOnlyElements: paraBgOnlyEls } = extractInlineBadgeShapes(child, slideRect, rect);\n        const {\n          shapes: paraDecorationShapes,\n          elements: paraDecorationEls\n        } = extractInlineDecorationShapes(child, slideRect);\n        if (paraDecorationShapes.length > 0)\n          elements.push(...paraDecorationShapes);\n        const paraLeadingOffset = computeLeadingOffset(\n          paraBadgeShapes,\n          rect,\n          slideRect\n        );\n        let inlineImgXOffset = 0;\n        let inlineImgYOffset = 0;\n        {\n          let firstNonEmojiImg = null;\n          let seenBrAfterImg = false;\n          for (const node of Array.from(child.childNodes)) {\n            if (node.nodeType === Node.ELEMENT_NODE) {\n              const en = node;\n              const enTag = en.tagName.toLowerCase();\n              if (enTag === "img") {\n                const ie = en;\n                if (!isEmojiImg(ie)) {\n                  if (!firstNonEmojiImg) firstNonEmojiImg = ie;\n                }\n                continue;\n              }\n              if (enTag === "br" && firstNonEmojiImg && !seenBrAfterImg) {\n                seenBrAfterImg = true;\n                continue;\n              }\n              break;\n            } else if (node.nodeType === Node.TEXT_NODE) {\n              if ((node.textContent ?? "").trim() !== "") break;\n            }\n          }\n          if (firstNonEmojiImg) {\n            const imgR = firstNonEmojiImg.getBoundingClientRect();\n            if (seenBrAfterImg) {\n              inlineImgYOffset = imgR.bottom - rect.top;\n            } else {\n              inlineImgXOffset = imgR.right - rect.left;\n              const parsedLH = parseFloat(style.lineHeight);\n              const lineHeight = !isNaN(parsedLH) && parsedLH > 0 ? parsedLH : (parseFloat(style.fontSize) || 16) * 1.5;\n              inlineImgYOffset = Math.max(0, imgR.bottom - rect.top - lineHeight);\n            }\n          }\n        }\n        const paraBadgeSet = paraBadgeEls.length > 0 ? new Set(paraBadgeEls) : false;\n        const paraBgOnlySet = paraBgOnlyEls.length > 0 || paraDecorationEls.length > 0 ? /* @__PURE__ */ new Set([...paraBgOnlyEls, ...paraDecorationEls]) : false;\n        if (paraBadgeEls.length > 0) {\n          const textSegments = extractInlineTextSegments(\n            child,\n            slideRect,\n            new Set(paraBadgeEls),\n            new Map(\n              paraBadgeEls.map((element, index) => [\n                element,\n                paraBadgeShapes[index]\n              ])\n            )\n          );\n          const hasMeasuredText = textSegments.some(\n            (segment) => segment.type === "paragraph"\n          );\n          if (hasMeasuredText) {\n            elements.push(...textSegments);\n            elements.push(...extractNestedImages(child, slideRect));\n            return elements;\n          }\n          elements.push(...paraBadgeShapes);\n        }\n        {\n          const extractedRuns = extractTextRuns(\n            child,\n            paraBadgeSet,\n            paraBgOnlySet\n          );\n          const measuredLayout = paraBadgeEls.length === 0 && paraDecorationEls.length === 0 ? extractMeasuredParagraphLayout(\n            child,\n            extractedRuns,\n            slideRect\n          ) : void 0;\n          const runs = measuredLayout?.runs ?? extractedRuns;\n          if (runs.some((r) => !r.breakLine && r.text.trim() !== "")) {\n            elements.push({\n              type: "paragraph",\n              runs,\n              ...base,\n              x: base.x + paraLeadingOffset + inlineImgXOffset,\n              y: base.y + inlineImgYOffset,\n              width: Math.max(10, base.width - paraLeadingOffset - inlineImgXOffset),\n              height: Math.max(10, base.height - inlineImgYOffset),\n              style: extractTextStyle(style),\n              ...measuredLayout && measuredLayout.lines.length > 1 ? { measuredLines: measuredLayout.lines } : {}\n            });\n          }\n        }\n        elements.push(...extractNestedImages(child, slideRect));\n      } else if (tag === "ul" || tag === "ol") {\n        const liChildren = Array.from(child.children).filter(\n          (c) => c.tagName.toLowerCase() === "li"\n        );\n        const hasEmbeddedImage = liChildren.some(\n          (li) => Array.from(li.querySelectorAll("img")).some(\n            (img) => !isEmojiImg(img)\n          )\n        );\n        const hasEmbeddedCode = liChildren.some(\n          (li) => Array.from(li.children).some((c) => {\n            const t = c.tagName.toLowerCase();\n            return t === "pre" || t === "marp-pre";\n          })\n        );\n        const olStart = tag === "ol" ? child.start || 1 : 1;\n        const olHasExplicitStart = tag === "ol" && child.hasAttribute("start");\n        if (!hasEmbeddedImage && !hasEmbeddedCode) {\n          const liBadgeSets = /* @__PURE__ */ new Map();\n          const liBgOnlySets = /* @__PURE__ */ new Map();\n          const liLeadingOffsetMap = /* @__PURE__ */ new Map();\n          const liBadgeEntries = /* @__PURE__ */ new Map();\n          for (const li of liChildren) {\n            const liRect = li.getBoundingClientRect();\n            const { shapes: liBadgeShapes, elements: liBadgeEls, bgOnlyElements: liBgOnlyEls } = extractInlineBadgeShapes(li, slideRect, liRect);\n            const {\n              shapes: liDecorationShapes,\n              elements: liDecorationEls\n            } = extractInlineDecorationShapes(li, slideRect);\n            if (liDecorationShapes.length > 0)\n              elements.push(...liDecorationShapes);\n            const directBadgeIndexes = liBadgeEls.map((element, index) => ({ element, index })).filter(({ element }) => element.closest("li") === li);\n            const directBadgeElements = directBadgeIndexes.map(\n              ({ element }) => element\n            );\n            const directBadgeShapes = directBadgeIndexes.map(\n              ({ index }) => liBadgeShapes[index]\n            );\n            const liLeadingOffset = computeLeadingOffset(\n              directBadgeShapes,\n              liRect,\n              slideRect\n            );\n            if (directBadgeShapes.length > 0) {\n              liBadgeEntries.set(\n                li,\n                directBadgeElements.map((element, index) => ({\n                  element,\n                  shape: directBadgeShapes[index]\n                }))\n              );\n            }\n            if (directBadgeElements.length > 0) {\n              liBadgeSets.set(li, new Set(directBadgeElements));\n            }\n            if (liBgOnlyEls.length > 0 || liDecorationEls.length > 0) {\n              liBgOnlySets.set(\n                li,\n                /* @__PURE__ */ new Set(\n                  [\n                    ...liBgOnlyEls.filter(\n                      (element) => element.closest("li") === li\n                    ),\n                    ...liDecorationEls.filter(\n                      (element) => element.closest("li") === li\n                    )\n                  ]\n                )\n              );\n            }\n            if (liLeadingOffset > 0) {\n              liLeadingOffsetMap.set(li, liLeadingOffset);\n            }\n          }\n          if (liBadgeEntries.size === 0) {\n            elements.push({\n              type: "list",\n              ordered: tag === "ol",\n              ...olHasExplicitStart ? { startNumber: olStart } : {},\n              listStyleType: style.listStyleType || void 0,\n              items: extractListItems(child),\n              ...base,\n              availableWidth: extractAvailableFlowWidth(child, rect),\n              style: extractTextStyle(style)\n            });\n          } else {\n            let orderedIndex = olStart;\n            let pendingItems = [];\n            let pendingTop = 0;\n            let pendingBottom = 0;\n            const flushPending = () => {\n              if (pendingItems.length === 0) return;\n              elements.push({\n                type: "list",\n                ordered: tag === "ol",\n                ...olHasExplicitStart || orderedIndex > 1 ? { startNumber: orderedIndex } : {},\n                listStyleType: style.listStyleType || void 0,\n                items: pendingItems,\n                x: base.x,\n                y: pendingTop,\n                width: base.width,\n                height: Math.max(10, pendingBottom - pendingTop),\n                availableWidth: extractAvailableFlowWidth(child, rect),\n                style: extractTextStyle(style)\n              });\n              orderedIndex += pendingItems.length;\n              pendingItems = [];\n            };\n            for (const li of liChildren) {\n              const liRect = li.getBoundingClientRect();\n              const badgeEntries = liBadgeEntries.get(li);\n              if (!badgeEntries) {\n                const liItems = extractListItemEl(li);\n                if (pendingItems.length === 0) {\n                  pendingTop = liRect.top - slideRect.top;\n                }\n                pendingBottom = liRect.bottom - slideRect.top;\n                pendingItems.push(...liItems);\n                continue;\n              }\n              flushPending();\n              const flow = extractInlineTextSegments(\n                li,\n                slideRect,\n                new Set(badgeEntries.map(({ element }) => element)),\n                new Map(\n                  badgeEntries.map(({ element, shape }) => [element, shape])\n                )\n              );\n              if (!flow.some((element) => element.type === "paragraph")) {\n                const fallbackItems = extractListItemEl(\n                  li,\n                  0,\n                  liBadgeSets.get(li) ?? false,\n                  liBgOnlySets.get(li) ?? false,\n                  liLeadingOffsetMap.get(li)\n                );\n                const fallbackList = () => {\n                  if (fallbackItems.length === 0) return;\n                  const badgeLeft = Math.min(\n                    ...badgeEntries.map(({ shape }) => shape.x)\n                  );\n                  elements.push({\n                    type: "list",\n                    ordered: tag === "ol",\n                    ...olHasExplicitStart || orderedIndex > 1 ? { startNumber: orderedIndex } : {},\n                    listStyleType: style.listStyleType || void 0,\n                    items: fallbackItems,\n                    x: base.x,\n                    y: liRect.top - slideRect.top,\n                    width: Math.max(\n                      10,\n                      Math.min(\n                        liRect.width,\n                        badgeLeft - (liRect.left - slideRect.left) - 2\n                      )\n                    ),\n                    height: liRect.height,\n                    availableWidth: extractAvailableFlowWidth(child, rect),\n                    style: extractTextStyle(style)\n                  });\n                  orderedIndex += 1;\n                };\n                if ((liLeadingOffsetMap.get(li) ?? 0) > 0) {\n                  elements.push(...badgeEntries.map(({ shape }) => shape));\n                  fallbackList();\n                } else {\n                  fallbackList();\n                  elements.push(...badgeEntries.map(({ shape }) => shape));\n                }\n                continue;\n              }\n              let listPrefixSegments = [];\n              let listPrefixEmitted = false;\n              const emitPrefixList = () => {\n                if (listPrefixEmitted || listPrefixSegments.length === 0)\n                  return;\n                const badgeLeft = Math.min(\n                  ...badgeEntries.map(({ shape }) => shape.x)\n                );\n                const listStyle = extractTextStyle(style);\n                const firstSegment = listPrefixSegments[0];\n                elements.push({\n                  type: "list",\n                  ordered: tag === "ol",\n                  ...tag === "ol" ? { startNumber: orderedIndex } : {},\n                  listStyleType: style.listStyleType || void 0,\n                  items: [{\n                    text: "\\xA0",\n                    level: 0,\n                    listStyleType: style.listStyleType || void 0,\n                    runs: [{\n                      text: "\\xA0",\n                      color: listStyle.color,\n                      fontSize: listStyle.fontSize,\n                      fontFamily: listStyle.fontFamily,\n                      bold: listStyle.fontWeight >= 600\n                    }]\n                  }],\n                  x: base.x,\n                  y: firstSegment.y,\n                  width: Math.max(\n                    10,\n                    firstSegment.x - base.x - 2 + 40\n                  ),\n                  height: Math.max(firstSegment.height, liRect.height),\n                  style: listStyle\n                });\n                elements.push(\n                  ...listPrefixSegments.map((segment) => ({\n                    ...segment,\n                    width: Math.max(\n                      10,\n                      badgeLeft - segment.x - 2\n                    )\n                  }))\n                );\n                listPrefixEmitted = true;\n                orderedIndex++;\n              };\n              for (const flowElement of flow) {\n                if (flowElement.type === "container") {\n                  emitPrefixList();\n                  elements.push(flowElement);\n                } else if (flowElement.type === "paragraph" && !listPrefixEmitted) {\n                  listPrefixSegments.push(flowElement);\n                } else {\n                  elements.push(flowElement);\n                }\n              }\n              emitPrefixList();\n            }\n            flushPending();\n          }\n          elements.push(...extractNestedImages(child, slideRect));\n          for (const li of liChildren) {\n            for (const preChild of Array.from(li.children)) {\n              const preChildTag = preChild.tagName.toLowerCase();\n              if (preChildTag !== "pre" && preChildTag !== "marp-pre") continue;\n              if (preChild.querySelector("svg")) continue;\n              const preRect = preChild.getBoundingClientRect();\n              if (preRect.width <= 0 || preRect.height <= 0) continue;\n              const preStyle = getComputedStyle(preChild);\n              const codeEl = preChild.querySelector("code");\n              const codeTarget = codeEl ?? preChild;\n              const preBg = preStyle.backgroundColor;\n              const codeBg = preBg && preBg !== "transparent" && preBg !== "rgba(0, 0, 0, 0)" && preBg !== "rgba(0,0,0,0)" ? preBg : codeEl ? getComputedStyle(codeEl).backgroundColor : preBg;\n              const preCodeRuns = extractCodeRuns(codeTarget);\n              const preScaleFactor = computeAutoScaleFactor(\n                preCodeRuns,\n                preStyle,\n                preRect.height,\n                computeElementTransformScale(preChild)\n              );\n              const preBaseStyle = extractTextStyle(preStyle);\n              const preAdjustedFs = applyAutoScale(preCodeRuns, preBaseStyle.fontSize, preScaleFactor);\n              elements.push({\n                type: "code",\n                text: codeTarget.textContent ?? "",\n                language: codeEl?.className?.replace("language-", "") ?? "",\n                runs: preCodeRuns,\n                x: preRect.left - slideRect.left,\n                y: preRect.top - slideRect.top,\n                width: preRect.width,\n                height: preRect.height,\n                style: {\n                  ...preBaseStyle,\n                  fontSize: preAdjustedFs,\n                  backgroundColor: codeBg\n                }\n              });\n            }\n          }\n        } else {\n          let olRunningNumber = olStart;\n          let pendingItems = [];\n          let pendingTop = -1;\n          let pendingBottom = -1;\n          const flushPending = () => {\n            if (pendingItems.length === 0) return;\n            const chunkStart = olRunningNumber;\n            elements.push({\n              type: "list",\n              ordered: tag === "ol",\n              ...olHasExplicitStart || chunkStart > 1 ? { startNumber: chunkStart } : {},\n              listStyleType: style.listStyleType || void 0,\n              items: pendingItems,\n              x: base.x,\n              y: pendingTop,\n              width: base.width,\n              height: Math.max(10, pendingBottom - pendingTop),\n              style: extractTextStyle(style)\n            });\n            olRunningNumber += pendingItems.length;\n            pendingItems = [];\n            pendingTop = -1;\n            pendingBottom = -1;\n          };\n          for (const li of liChildren) {\n            const liImages = Array.from(li.querySelectorAll("img")).filter(\n              (img) => !isEmojiImg(img)\n            );\n            const liRect = li.getBoundingClientRect();\n            const liY = liRect.top - slideRect.top;\n            const liBottom = liRect.bottom - slideRect.top;\n            const { shapes: liSplitBadgeShapes, elements: liSplitBadgeEls, bgOnlyElements: liSplitBgOnlyEls } = extractInlineBadgeShapes(li, slideRect, liRect);\n            const liLeadingOffset = computeLeadingOffset(\n              liSplitBadgeShapes,\n              liRect,\n              slideRect\n            );\n            const liSplitSkipBadges = liSplitBadgeEls.length > 0 ? new Set(liSplitBadgeEls) : false;\n            const liSplitStripBadges = liSplitBgOnlyEls.length > 0 ? new Set(liSplitBgOnlyEls) : false;\n            if (liSplitBadgeShapes.length > 0) elements.push(...liSplitBadgeShapes);\n            if (liImages.length === 0) {\n              const liCodeChildren = Array.from(li.children).filter((c) => {\n                const t = c.tagName.toLowerCase();\n                return (t === "pre" || t === "marp-pre") && !c.querySelector("svg");\n              });\n              const liItems = extractListItemEl(\n                li,\n                0,\n                liSplitSkipBadges,\n                liSplitStripBadges,\n                liLeadingOffset\n              );\n              if (liCodeChildren.length === 0) {\n                if (pendingTop < 0) pendingTop = liY;\n                pendingBottom = liBottom;\n                pendingItems.push(...liItems);\n              } else {\n                if (pendingTop < 0) pendingTop = liY;\n                const firstCodeRect = liCodeChildren[0].getBoundingClientRect();\n                pendingBottom = firstCodeRect.top - slideRect.top;\n                pendingItems.push(...liItems);\n                flushPending();\n                for (const preChild of liCodeChildren) {\n                  const preRect = preChild.getBoundingClientRect();\n                  if (preRect.width <= 0 || preRect.height <= 0) continue;\n                  const preStyle = getComputedStyle(preChild);\n                  const codeEl = preChild.querySelector("code");\n                  const codeTarget = codeEl ?? preChild;\n                  const preBg = preStyle.backgroundColor;\n                  const codeBg = preBg && preBg !== "transparent" && preBg !== "rgba(0, 0, 0, 0)" && preBg !== "rgba(0,0,0,0)" ? preBg : codeEl ? getComputedStyle(codeEl).backgroundColor : preBg;\n                  const preCodeRuns = extractCodeRuns(codeTarget);\n                  const preScaleFactor = computeAutoScaleFactor(\n                    preCodeRuns,\n                    preStyle,\n                    preRect.height,\n                    computeElementTransformScale(preChild)\n                  );\n                  const preBaseStyle = extractTextStyle(preStyle);\n                  const preAdjustedFs = applyAutoScale(preCodeRuns, preBaseStyle.fontSize, preScaleFactor);\n                  elements.push({\n                    type: "code",\n                    text: codeTarget.textContent ?? "",\n                    language: codeEl?.className?.replace("language-", "") ?? "",\n                    runs: preCodeRuns,\n                    x: preRect.left - slideRect.left,\n                    y: preRect.top - slideRect.top,\n                    width: preRect.width,\n                    height: preRect.height,\n                    style: {\n                      ...preBaseStyle,\n                      fontSize: preAdjustedFs,\n                      backgroundColor: codeBg\n                    }\n                  });\n                }\n              }\n            } else {\n              const liItems = extractListItemEl(\n                li,\n                0,\n                liSplitSkipBadges,\n                liSplitStripBadges,\n                liLeadingOffset\n              );\n              if (liItems.length > 0) {\n                if (pendingTop < 0) pendingTop = liY;\n                pendingBottom = liBottom;\n                pendingItems.push(...liItems);\n              }\n              flushPending();\n              for (const img of liImages) {\n                const imgRect = img.getBoundingClientRect();\n                const imgFilter = getComputedStyle(img).filter && getComputedStyle(img).filter !== "none" ? getComputedStyle(img).filter : void 0;\n                elements.push({\n                  type: "image",\n                  src: img.src,\n                  naturalWidth: img.naturalWidth,\n                  naturalHeight: img.naturalHeight,\n                  x: imgRect.left - slideRect.left,\n                  y: imgRect.top - slideRect.top,\n                  width: imgRect.width,\n                  height: imgRect.height,\n                  ...imgFilter ? { cssFilter: imgFilter, pageX: imgRect.left, pageY: imgRect.top } : {}\n                });\n              }\n            }\n          }\n          flushPending();\n        }\n      } else if (tag === "table") {\n        const { rows: tableRows, colWidths } = extractTableData(child);\n        elements.push({\n          type: "table",\n          rows: tableRows,\n          ...colWidths.length > 0 ? { colWidths } : {},\n          ...base,\n          style: extractTextStyle(style)\n        });\n        elements.push(...extractNestedImages(child, slideRect));\n      } else if (tag === "pre" || tag === "marp-pre") {\n        const innerSvg = child.querySelector("svg");\n        if (innerSvg) {\n          try {\n            const svgStr = new XMLSerializer().serializeToString(innerSvg);\n            const b64 = btoa(unescape(encodeURIComponent(svgStr)));\n            const dataUrl = `data:image/svg+xml;base64,${b64}`;\n            elements.push({\n              type: "image",\n              src: dataUrl,\n              naturalWidth: base.width,\n              naturalHeight: base.height,\n              ...base,\n              // Request rasterization: Mermaid SVGs may use <foreignObject>\n              // for text labels which PowerPoint cannot render from SVG data.\n              // pageX/pageY are intentionally omitted: rasterizeSlideTargets\n              // computes the absolute clip from the slide-relative x/y after\n              // navigating to the correct slide (avoids stale bespoke-transform\n              // coordinates).\n              rasterize: true\n            });\n          } catch {\n            const code = child.querySelector("code");\n            const codeTarget = code ?? child;\n            const preBgCatch = style.backgroundColor;\n            const codeBgCatch = preBgCatch && preBgCatch !== "transparent" && preBgCatch !== "rgba(0, 0, 0, 0)" && preBgCatch !== "rgba(0,0,0,0)" ? preBgCatch : code ? getComputedStyle(code).backgroundColor : preBgCatch;\n            const codeRunsCatch = extractCodeRuns(codeTarget);\n            const scaleFactorCatch = computeAutoScaleFactor(\n              codeRunsCatch,\n              style,\n              rect.height,\n              computeElementTransformScale(child)\n            );\n            const baseStyleCatch = extractTextStyle(style);\n            const adjustedFsCatch = applyAutoScale(codeRunsCatch, baseStyleCatch.fontSize, scaleFactorCatch);\n            elements.push({\n              type: "code",\n              text: codeTarget.textContent ?? "",\n              language: code?.className?.replace("language-", "") ?? "",\n              runs: codeRunsCatch,\n              ...base,\n              style: {\n                ...baseStyleCatch,\n                fontSize: adjustedFsCatch,\n                backgroundColor: codeBgCatch\n              }\n            });\n          }\n        } else {\n          const code = child.querySelector("code");\n          const codeTarget = code ?? child;\n          const preBgStd = style.backgroundColor;\n          const codeBgStd = preBgStd && preBgStd !== "transparent" && preBgStd !== "rgba(0, 0, 0, 0)" && preBgStd !== "rgba(0,0,0,0)" ? preBgStd : code ? getComputedStyle(code).backgroundColor : preBgStd;\n          const codeRuns = extractCodeRuns(codeTarget);\n          const scaleFactor = computeAutoScaleFactor(\n            codeRuns,\n            style,\n            rect.height,\n            computeElementTransformScale(child)\n          );\n          const baseStyle = extractTextStyle(style);\n          const adjustedFontSize = applyAutoScale(codeRuns, baseStyle.fontSize, scaleFactor);\n          elements.push({\n            type: "code",\n            text: codeTarget.textContent ?? "",\n            language: code?.className?.replace("language-", "") ?? "",\n            runs: codeRuns,\n            ...base,\n            style: {\n              ...baseStyle,\n              fontSize: adjustedFontSize,\n              backgroundColor: codeBgStd,\n              ...parseFloat(style.paddingTop) > 0 ? { paddingTop: parseFloat(style.paddingTop) } : {},\n              ...parseFloat(style.paddingRight) > 0 ? { paddingRight: parseFloat(style.paddingRight) } : {},\n              ...parseFloat(style.paddingBottom) > 0 ? { paddingBottom: parseFloat(style.paddingBottom) } : {},\n              ...parseFloat(style.paddingLeft) > 0 ? { paddingLeft: parseFloat(style.paddingLeft) } : {}\n            }\n          });\n        }\n      } else if (tag === "img") {\n        const img = child;\n        if (img.classList?.contains("emoji") || img.src?.includes("twemoji") || img.src?.includes("/emoji/"))\n          return elements;\n        const imgFilter = style.filter && style.filter !== "none" ? style.filter : void 0;\n        const clipped = resolveClippedRect(img, rect, slideRect);\n        if (clipped.width <= 0 || clipped.height <= 0) return elements;\n        elements.push({\n          type: "image",\n          src: img.src,\n          naturalWidth: img.naturalWidth,\n          naturalHeight: img.naturalHeight,\n          x: clipped.left - slideRect.left,\n          y: clipped.top - slideRect.top,\n          width: clipped.width,\n          height: clipped.height,\n          // Store page-absolute coords when cssFilter is set so the export\n          // tool can screenshot the rendered (filtered) region via Puppeteer.\n          ...clipped.isClipped ? { rasterize: true } : imgFilter ? { cssFilter: imgFilter, pageX: rect.left, pageY: rect.top } : {}\n        });\n      } else if (tag === "blockquote") {\n        const borderWidth = parseFloat(style.borderLeftWidth) || 0;\n        const borderColor = style.borderLeftColor;\n        const paddingTop = parseFloat(style.paddingTop) || 0;\n        const paddingRight = parseFloat(style.paddingRight) || 0;\n        const paddingBottom = parseFloat(style.paddingBottom) || 0;\n        const paddingLeft = parseFloat(style.paddingLeft) || 0;\n        const { shapes: bqBadgeShapes, elements: bqBadgeEls, bgOnlyElements: bqBgOnlyEls } = extractInlineBadgeShapes(child, slideRect, rect);\n        const bqLeadingOffset = computeLeadingOffset(\n          bqBadgeShapes,\n          rect,\n          slideRect\n        );\n        if (bqBadgeShapes.length > 0) elements.push(...bqBadgeShapes);\n        const bqBadgeSet = bqBadgeEls.length > 0 ? new Set(bqBadgeEls) : false;\n        const bqBgOnlySet = bqBgOnlyEls.length > 0 ? new Set(bqBgOnlyEls) : false;\n        const directListEls = Array.from(child.children).filter((c) => {\n          const t = c.tagName.toLowerCase();\n          return t === "ul" || t === "ol";\n        });\n        const directListSet = directListEls.length > 0 ? new Set(directListEls) : false;\n        elements.push({\n          type: "blockquote",\n          runs: extractTextRuns(child, bqBadgeSet, bqBgOnlySet, directListSet),\n          ...base,\n          x: base.x + bqLeadingOffset,\n          width: Math.max(10, base.width - bqLeadingOffset),\n          style: {\n            ...extractTextStyle(style),\n            ...paddingTop || paddingRight || paddingBottom || paddingLeft ? { paddingTop, paddingRight, paddingBottom, paddingLeft } : {}\n          },\n          ...borderWidth > 0 ? { borderLeft: { width: borderWidth, color: borderColor } } : {}\n        });\n        for (const listEl of directListEls) {\n          const listRect = listEl.getBoundingClientRect();\n          if (listRect.width <= 0 || listRect.height <= 0) continue;\n          const listStyle = getComputedStyle(listEl);\n          const listTag = listEl.tagName.toLowerCase();\n          const listOrdered = listTag === "ol";\n          const olStart = listOrdered ? listEl.start || 1 : 1;\n          const olHasExplicitStart = listOrdered && listEl.hasAttribute("start");\n          elements.push({\n            type: "list",\n            ordered: listOrdered,\n            ...olHasExplicitStart ? { startNumber: olStart } : {},\n            listStyleType: listStyle.listStyleType || void 0,\n            items: extractListItems(listEl, 0),\n            x: listRect.left - slideRect.left,\n            y: listRect.top - slideRect.top,\n            width: listRect.width,\n            height: listRect.height,\n            style: extractTextStyle(listStyle)\n          });\n        }\n        elements.push(...extractNestedImages(child, slideRect));\n      } else if (tag === "svg") {\n        try {\n          const svgStr = new XMLSerializer().serializeToString(child);\n          const b64 = btoa(unescape(encodeURIComponent(svgStr)));\n          const dataUrl = `data:image/svg+xml;base64,${b64}`;\n          const hasForeignObject = child.querySelector("foreignObject") !== null;\n          const clipped = resolveClippedRect(child, rect, slideRect);\n          if (clipped.width <= 0 || clipped.height <= 0) return elements;\n          elements.push({\n            type: "image",\n            src: dataUrl,\n            naturalWidth: clipped.width,\n            naturalHeight: clipped.height,\n            x: clipped.left - slideRect.left,\n            y: clipped.top - slideRect.top,\n            width: clipped.width,\n            height: clipped.height,\n            ...hasForeignObject || clipped.isClipped ? { rasterize: true } : {}\n          });\n        } catch {\n        }\n      } else if (tag === "header" || tag === "footer") {\n        const { shapes: hfBadgeShapes, elements: hfBadgeEls, bgOnlyElements: hfBgOnlyEls } = extractInlineBadgeShapes(child, slideRect, rect);\n        const hfLeadingOffset = computeLeadingOffset(\n          hfBadgeShapes,\n          rect,\n          slideRect\n        );\n        if (hfBadgeShapes.length > 0) elements.push(...hfBadgeShapes);\n        const hfBadgeSet = hfBadgeEls.length > 0 ? new Set(hfBadgeEls) : false;\n        const hfBgOnlySet = hfBgOnlyEls.length > 0 ? new Set(hfBgOnlyEls) : false;\n        const hfWidth = Math.max(\n          base.width - hfLeadingOffset,\n          slideRect.width - base.x - hfLeadingOffset\n        );\n        elements.push({\n          type: tag,\n          runs: extractTextRuns(child, hfBadgeSet, hfBgOnlySet),\n          ...base,\n          x: base.x + hfLeadingOffset,\n          width: hfWidth,\n          style: extractTextStyle(style)\n        });\n        elements.push(...extractNestedImages(child, slideRect));\n      } else {\n        const borders = extractCssBorders(style);\n        const borderSides = [\n          borders.top,\n          borders.right,\n          borders.bottom,\n          borders.left\n        ].filter((border) => border !== void 0);\n        const hasBorder = borderSides.length > 0;\n        const borderRadius = parseFloat(style.borderRadius) || 0;\n        const borderLeftWidth = borders.left?.width ?? 0;\n        const hasBorderLeft = borderLeftWidth > 0 && borderSides.length === 1;\n        const borderBottomWidth = borders.bottom?.width ?? 0;\n        const hasBorderBottom = borderBottomWidth > 0 && borderSides.length === 1;\n        const boxShadow = style.boxShadow;\n        const hasBoxShadow = !!boxShadow && boxShadow !== "none";\n        const hasBackground = !!style.backgroundColor && style.backgroundColor !== "transparent" && style.backgroundColor !== "rgba(0, 0, 0, 0)" && !style.backgroundColor.match(\n          /rgba\\(\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*\\d+\\s*,\\s*0(?:\\.0+)?\\s*\\)/\n        );\n        const blockChildren = walkElements(child, slideRect);\n        const containerStyle = {\n          backgroundColor: style.backgroundColor,\n          ...hasBorder ? {\n            borders,\n            borderWidth: borders.top?.width ?? borders.right?.width ?? borders.bottom?.width ?? borders.left.width,\n            borderColor: borders.top?.color ?? borders.right?.color ?? borders.bottom?.color ?? borders.left.color,\n            borderStyle: borders.top?.style ?? borders.right?.style ?? borders.bottom?.style ?? borders.left?.style ?? "solid"\n          } : {},\n          ...borderRadius > 0 ? { borderRadius } : {},\n          ...hasBorderLeft ? {\n            borderLeft: {\n              width: borderLeftWidth,\n              color: borders.left.color\n            }\n          } : {},\n          ...hasBorderBottom ? {\n            borderBottom: {\n              width: borderBottomWidth,\n              color: borders.bottom.color,\n              style: borders.bottom.style\n            }\n          } : {},\n          ...hasBoxShadow ? { boxShadow: true } : {}\n        };\n        const containerIsFlexOrGrid = /^(flex|inline-flex|grid|inline-grid)/.test(\n          style.display\n        );\n        const blockChildrenAllInlineLevel = !containerIsFlexOrGrid && Array.from(child.children).every((node) => {\n          const nodeStyle = getComputedStyle(node);\n          return nodeStyle.display === "inline" || /^inline-/.test(nodeStyle.display);\n        });\n        if (blockChildren.length > 0) {\n          const directElements = Array.from(child.children);\n          const directOutputs = /* @__PURE__ */ new Map();\n          const unownedOutputs = [];\n          for (const output of blockChildren) {\n            const owner = directElements.find((directElement) => {\n              const directPath = sourceRefFor(directElement).domPath;\n              const outputPath = output.source?.domPath;\n              return !!outputPath && (outputPath === directPath || outputPath.startsWith(`${directPath}/`));\n            });\n            if (owner) {\n              const outputs = directOutputs.get(owner) ?? [];\n              outputs.push(output);\n              directOutputs.set(owner, outputs);\n            } else {\n              unownedOutputs.push(output);\n            }\n          }\n          const orderedChildren = [];\n          for (const node of Array.from(child.childNodes)) {\n            if (node.nodeType === Node.ELEMENT_NODE) {\n              const directTag = node.tagName.toLowerCase();\n              if (directTag === "svg" || directTag === "mjx-container") {\n                continue;\n              }\n              const ownedOutputs = directOutputs.get(node);\n              if (ownedOutputs) {\n                orderedChildren.push(...ownedOutputs);\n              } else {\n                const inlineSegments = extractInlineTextSegments(\n                  node,\n                  slideRect,\n                  /* @__PURE__ */ new Set()\n                );\n                if (inlineSegments.length > 0) {\n                  orderedChildren.push(...inlineSegments);\n                } else {\n                  const inlineElement = node;\n                  const inlineStyle = getComputedStyle(inlineElement);\n                  if (inlineStyle.display === "inline") {\n                    const inlineRect = inlineElement.getBoundingClientRect();\n                    const inlineRuns = extractTextRuns(inlineElement);\n                    if (inlineRuns.some(\n                      (run) => !run.breakLine && run.text.trim() !== ""\n                    )) {\n                      orderedChildren.push({\n                        type: "paragraph",\n                        source: sourceRefFor(inlineElement),\n                        x: inlineRect.left - slideRect.left,\n                        y: inlineRect.top - slideRect.top,\n                        width: Math.max(10, inlineRect.width),\n                        height: Math.max(10, inlineRect.height),\n                        runs: inlineRuns,\n                        style: extractTextStyle(inlineStyle)\n                      });\n                    }\n                  }\n                }\n              }\n            } else if (node.nodeType === Node.TEXT_NODE) {\n              const text = node.textContent?.trim() ?? "";\n              if (!text) continue;\n              const range = document.createRange();\n              range.selectNodeContents(node);\n              const hasMeasuredRange = typeof range.getBoundingClientRect === "function";\n              if (!hasMeasuredRange && !containerIsFlexOrGrid && !blockChildrenAllInlineLevel)\n                continue;\n              const textRect = hasMeasuredRange ? range.getBoundingClientRect() : rect;\n              if (textRect.width <= 0 || textRect.height <= 0) continue;\n              const fallbackX = containerIsFlexOrGrid ? computeShallowFlexOffset(child, rect, slideRect, style) : 0;\n              orderedChildren.push({\n                type: "paragraph",\n                source: {\n                  ...sourceRefFor(child),\n                  id: `${sourceRefFor(child).id}:text`\n                },\n                x: hasMeasuredRange ? textRect.left - slideRect.left : base.x + fallbackX,\n                y: textRect.top - slideRect.top,\n                width: hasMeasuredRange ? textRect.width : Math.max(10, base.width - fallbackX),\n                height: textRect.height,\n                runs: [{ text }],\n                style: extractTextStyle(style, false)\n              });\n            }\n          }\n          orderedChildren.push(...unownedOutputs);\n          elements.push({\n            type: "container",\n            children: orderedChildren,\n            ...base,\n            style: containerStyle\n          });\n        } else {\n          if (hasBackground || hasBorder || hasBorderLeft || hasBorderBottom || hasBoxShadow) {\n            elements.push({\n              type: "container",\n              children: [],\n              ...base,\n              style: containerStyle\n            });\n          }\n          const runs = extractTextRuns(child);\n          if (hasBackground) {\n            const elBg = style.backgroundColor;\n            for (const r of runs) {\n              if (!r.breakLine && r.backgroundColor === elBg) {\n                r.backgroundColor = void 0;\n              }\n            }\n          }\n          if (runs.some((r) => !r.breakLine && r.text.trim() !== "")) {\n            const valign = style.alignItems === "center" || style.justifyContent === "center" || style.verticalAlign === "middle" ? "middle" : "top";\n            const paddingTop = parseFloat(style.paddingTop) || 0;\n            const paddingRight = parseFloat(style.paddingRight) || 0;\n            const paddingBottom = parseFloat(style.paddingBottom) || 0;\n            const paddingLeft = parseFloat(style.paddingLeft) || 0;\n            const parentRight = parentIsFlexOrGrid ? parent.getBoundingClientRect().right - slideRect.left : 0;\n            const emojiWidthOverride = (() => {\n              if (!parentIsFlexOrGrid) return void 0;\n              const hasEmoji = runs.some(\n                (r) => !r.breakLine && /\\p{Extended_Pictographic}/u.test(r.text)\n              );\n              if (!hasEmoji) return void 0;\n              let sib = child.nextElementSibling;\n              while (sib !== null) {\n                const sibStyle = getComputedStyle(sib);\n                if (sibStyle.display !== "none" && sibStyle.visibility !== "hidden") {\n                  const sibRect = sib.getBoundingClientRect();\n                  if (sibRect.width > 0 && sibRect.height > 0) return void 0;\n                }\n                sib = sib.nextElementSibling;\n              }\n              const extended = Math.max(base.width, parentRight - base.x);\n              return extended > base.width ? extended : void 0;\n            })();\n            const nowrapWidthOverride = (() => {\n              if (emojiWidthOverride !== void 0) return void 0;\n              if (!parentIsFlexOrGrid) return void 0;\n              if (style.whiteSpace !== "nowrap") return void 0;\n              const wantedWidth = base.width * 1.15;\n              const slideWidth = slideRect.width;\n              const maxWidth = Math.max(base.width, slideWidth - base.x);\n              const finalWidth = Math.min(wantedWidth, maxWidth);\n              return finalWidth > base.width ? finalWidth : void 0;\n            })();\n            const nowrapXShift = 0;\n            elements.push({\n              type: "paragraph",\n              runs,\n              ...base,\n              ...emojiWidthOverride !== void 0 ? { width: emojiWidthOverride } : nowrapWidthOverride !== void 0 ? { width: nowrapWidthOverride } : (\n                // Inline-only containers (e.g. display:inline-block badges)\n                // have tight-fitting widths from browser font metrics.\n                // PowerPoint fonts may render slightly wider, causing text to\n                // wrap.  Add a small slack (8 px) when the container has a\n                // visible background (badge/chip pattern) to absorb the\n                // font-metric variance.\n                hasBackground ? { width: base.width + 8 } : parentIsFlexOrGrid ? {\n                  // ADR-23: Non-background text in flex/grid child containers\n                  // (e.g. step-body divs, label divs) can wrap in PPTX when\n                  // DirectWrite font metrics are slightly wider than Chrome\'s\n                  // Skia. The wrapping causes the next absolutely-positioned\n                  // item to overlap. Add 16px slack capped at parent right edge.\n                  width: Math.min(\n                    base.width + 16,\n                    Math.max(base.width, parentRight - base.x)\n                  )\n                } : {}\n              ),\n              style: {\n                ...extractTextStyle(style),\n                ...paddingTop || paddingRight || paddingBottom || paddingLeft ? { paddingTop, paddingRight, paddingBottom, paddingLeft } : {}\n              },\n              valign\n            });\n          }\n          elements.push(...extractNestedImages(child, slideRect));\n        }\n      }\n      return elements;\n    }\n    function walkDisplayContents(container, effectiveParent, slideRect) {\n      const elements = [];\n      const containerStyle = getComputedStyle(container);\n      const textIsVisible = containerStyle.visibility !== "hidden";\n      for (const node of Array.from(container.childNodes)) {\n        if (node.nodeType === Node.ELEMENT_NODE) {\n          const child = node;\n          const walked = walkElement(effectiveParent, child, slideRect);\n          if (walked.length > 0) {\n            elements.push(...walked);\n            continue;\n          }\n          const childStyle = getComputedStyle(child);\n          const childRect = child.getBoundingClientRect();\n          if (childStyle.display !== "inline" || childStyle.visibility === "hidden" || childRect.width === 0 || childRect.height === 0)\n            continue;\n          const runs = extractTextRuns(child);\n          if (runs.some((run) => !run.breakLine && run.text.trim() !== "")) {\n            elements.push({\n              type: "paragraph",\n              runs,\n              x: childRect.left - slideRect.left,\n              y: childRect.top - slideRect.top,\n              width: childRect.width,\n              height: childRect.height,\n              style: extractTextStyle(childStyle)\n            });\n          }\n          continue;\n        }\n        if (node.nodeType !== Node.TEXT_NODE) continue;\n        if (!textIsVisible) continue;\n        const text = node.textContent ?? "";\n        if (text.trim() === "") continue;\n        const range = document.createRange();\n        range.selectNodeContents(node);\n        const rect = range.getBoundingClientRect();\n        if (rect.width === 0 || rect.height === 0) continue;\n        elements.push({\n          type: "paragraph",\n          source: sourceRefFor(container),\n          runs: [\n            {\n              text: text.trim(),\n              color: containerStyle.color,\n              fontSize: parseFloat(containerStyle.fontSize) || 16,\n              fontFamily: containerStyle.fontFamily,\n              bold: parseInt(containerStyle.fontWeight, 10) >= 600,\n              italic: containerStyle.fontStyle === "italic",\n              underline: containerStyle.textDecorationLine?.includes("underline"),\n              strikethrough: containerStyle.textDecorationLine?.includes("line-through")\n            }\n          ],\n          x: rect.left - slideRect.left,\n          y: rect.top - slideRect.top,\n          width: rect.width,\n          height: rect.height,\n          style: extractTextStyle(containerStyle)\n        });\n      }\n      return elements;\n    }\n    function walkElements(parent, slideRect) {\n      const elements = [];\n      for (const child of Array.from(parent.children)) {\n        elements.push(...walkElement(parent, child, slideRect));\n      }\n      return elements;\n    }\n    function walkSectionNodes(parent, slideRect) {\n      const elements = [];\n      const style = getComputedStyle(parent);\n      if (style.display === "none" || style.visibility === "hidden") return elements;\n      for (const node of Array.from(parent.childNodes)) {\n        if (node.nodeType === Node.ELEMENT_NODE) {\n          const child = node;\n          elements.push(...walkElement(parent, child, slideRect));\n          continue;\n        }\n        if (node.nodeType === Node.TEXT_NODE) {\n          const text = node.textContent?.trim() ?? "";\n          if (!text) continue;\n          const range = document.createRange();\n          range.selectNodeContents(node);\n          if (typeof range.getBoundingClientRect !== "function") continue;\n          const rect = range.getBoundingClientRect();\n          if (rect.width <= 0 || rect.height <= 0) continue;\n          elements.push({\n            type: "paragraph",\n            source: {\n              ...sourceRefFor(parent),\n              id: `${sourceRefFor(parent).id}:direct-text`\n            },\n            x: rect.left - slideRect.left,\n            y: rect.top - slideRect.top,\n            width: rect.width,\n            height: rect.height,\n            runs: [{\n              text,\n              color: style.color,\n              fontSize: parseFloat(style.fontSize) || 16,\n              fontFamily: style.fontFamily,\n              bold: parseInt(style.fontWeight, 10) >= 600,\n              italic: style.fontStyle === "italic",\n              underline: style.textDecorationLine?.includes("underline"),\n              strikethrough: style.textDecorationLine?.includes("line-through")\n            }],\n            style: extractTextStyle(style, false)\n          });\n        }\n      }\n      return elements;\n    }\n    const globalPseudoSignatures = /* @__PURE__ */ new Set();\n    function extractPseudoElements(host, slideRect, selectedPseudo) {\n      const shapes = [];\n      const pseudoElements = selectedPseudo ? [selectedPseudo] : ["::before", "::after"];\n      for (const pseudo of pseudoElements) {\n        const ps = getComputedStyle(host, pseudo);\n        const rawContent = ps.content;\n        if (!rawContent || rawContent === "none" || rawContent === "normal")\n          continue;\n        const isSection = host.tagName.toLowerCase() === "section";\n        if (!isSection && ps.position !== "absolute" && ps.position !== "fixed")\n          continue;\n        const stripped = rawContent.replace(/^["\']|["\']$/g, "").trim();\n        if (stripped === "" && isSection) {\n          const sectionClass = host.className?.trim() ?? "";\n          if (!sectionClass) continue;\n          const pgBg = ps.backgroundColor;\n          if (!pgBg || pgBg === "transparent" || pgBg === "rgba(0, 0, 0, 0)" || globalPseudoSignatures.has(`${pseudo}:${pgBg}`))\n            continue;\n        }\n        const bg = ps.backgroundColor;\n        const hasBackground = !!bg && bg !== "transparent" && bg !== "rgba(0, 0, 0, 0)";\n        const hasBackgroundImage = !!ps.backgroundImage && ps.backgroundImage !== "none";\n        const pseudoBorders = {\n          top: parseFloat(ps.borderTopWidth) || 0,\n          right: parseFloat(ps.borderRightWidth) || 0,\n          bottom: parseFloat(ps.borderBottomWidth) || 0,\n          left: parseFloat(ps.borderLeftWidth) || 0\n        };\n        const hasBorders = [\n          pseudoBorders.top,\n          pseudoBorders.right,\n          pseudoBorders.bottom,\n          pseudoBorders.left\n        ].some((width) => width > 0);\n        const hasTransform = !!ps.transform && ps.transform !== "none";\n        const hasVisualPaint = hasBackground || hasBackgroundImage || hasBorders || !!ps.boxShadow && ps.boxShadow !== "none" || stripped !== "";\n        if (!hasVisualPaint) continue;\n        const declaredWidth = parseFloat(ps.width) || 0;\n        const declaredHeight = parseFloat(ps.height) || 0;\n        const w = Math.max(\n          declaredWidth,\n          pseudoBorders.left + pseudoBorders.right\n        );\n        const h = Math.max(\n          declaredHeight,\n          pseudoBorders.top + pseudoBorders.bottom\n        );\n        if (w <= 0 || h <= 0) continue;\n        const position = ps.position;\n        const hostRect = host.getBoundingClientRect();\n        let x = hostRect.left - slideRect.left;\n        let y = hostRect.top - slideRect.top;\n        if (position === "absolute" || position === "fixed") {\n          const top = parseFloat(ps.top);\n          const left = parseFloat(ps.left);\n          const bottom = parseFloat(ps.bottom);\n          if (!isNaN(top)) y = hostRect.top - slideRect.top + top;\n          else if (!isNaN(bottom)) y = hostRect.top - slideRect.top + hostRect.height - bottom - h;\n          else y = hostRect.top - slideRect.top;\n          const right = parseFloat(ps.right);\n          if (!isNaN(left)) x = hostRect.left - slideRect.left + left;\n          else if (!isNaN(right) && !isSection)\n            x = hostRect.left - slideRect.left + hostRect.width - right - w;\n        }\n        const clippedLeft = Math.max(0, x);\n        const clippedTop = Math.max(0, y);\n        const clippedRight = Math.min(slideRect.width, x + w);\n        const clippedBottom = Math.min(slideRect.height, y + h);\n        if (clippedRight <= clippedLeft || clippedBottom <= clippedTop) continue;\n        const base = {\n          x: clippedLeft,\n          y: clippedTop,\n          width: clippedRight - clippedLeft,\n          height: clippedBottom - clippedTop,\n          ...extractStackingMetadata(ps),\n          ...extractStackingMetadata(getComputedStyle(host))\n        };\n        const source = sourceRefFor(host);\n        const pseudoNeedsRaster = hasBackgroundImage || hasBorders || hasTransform || stripped !== "" && !isSection || !!ps.boxShadow && ps.boxShadow !== "none";\n        shapes.push(\n          pseudoNeedsRaster ? {\n            type: "image",\n            source: {\n              ...source,\n              id: `${source.id}:${pseudo}`,\n              domPath: `${source.domPath}/${pseudo}`\n            },\n            src: "",\n            naturalWidth: base.width,\n            naturalHeight: base.height,\n            ...base,\n            rasterize: true\n          } : {\n            type: "container",\n            source: {\n              ...source,\n              id: `${source.id}:${pseudo}`,\n              domPath: `${source.domPath}/${pseudo}`\n            },\n            children: [],\n            ...base,\n            style: { backgroundColor: bg }\n          }\n        );\n      }\n      return shapes;\n    }\n    function extractDescendantPseudoElements(root2, slideRect) {\n      const elements = [];\n      for (const child of Array.from(root2.querySelectorAll("*"))) {\n        elements.push(...extractPseudoElements(child, slideRect, "::before"));\n        elements.push(...extractPseudoElements(child, slideRect, "::after"));\n      }\n      return elements;\n    }\n    function attachPseudoElements(elements, pseudoElements) {\n      const before = /* @__PURE__ */ new Map();\n      const after = /* @__PURE__ */ new Map();\n      const placed = /* @__PURE__ */ new Set();\n      for (const pseudo of pseudoElements) {\n        const sourceId = pseudo.source?.id;\n        if (!sourceId) continue;\n        const marker = sourceId.endsWith(":::after") ? "after" : "before";\n        const hostId = sourceId.replace(/:::(before|after)$/, "");\n        const map = marker === "after" ? after : before;\n        const bucket = map.get(hostId) ?? [];\n        bucket.push(pseudo);\n        map.set(hostId, bucket);\n      }\n      const nonContainerHosts = /* @__PURE__ */ new Set();\n      const collectNonContainerHosts = (element) => {\n        if (element.type !== "container" && element.source?.id) {\n          nonContainerHosts.add(element.source.id);\n        }\n        if (element.type === "container") {\n          element.children.forEach(collectNonContainerHosts);\n        }\n      };\n      elements.forEach(collectNonContainerHosts);\n      const visit = (element) => {\n        const sourceId = element.source?.id;\n        const canOwnPseudo = sourceId !== void 0 && (element.type !== "container" || !nonContainerHosts.has(sourceId)) && ![...before.get(sourceId) ?? [], ...after.get(sourceId) ?? []].some((pseudo) => placed.has(pseudo));\n        const attachedBefore = canOwnPseudo ? before.get(sourceId) ?? [] : [];\n        const attachedAfter = canOwnPseudo ? after.get(sourceId) ?? [] : [];\n        attachedBefore.forEach((pseudo) => placed.add(pseudo));\n        attachedAfter.forEach((pseudo) => placed.add(pseudo));\n        if (element.type === "container" && canOwnPseudo) {\n          element = {\n            ...element,\n            children: [\n              ...attachedBefore,\n              ...element.children.flatMap(visit),\n              ...attachedAfter\n            ]\n          };\n          return [element];\n        }\n        if (element.type === "container") {\n          element = {\n            ...element,\n            children: element.children.flatMap(visit)\n          };\n        }\n        return [...attachedBefore, element, ...attachedAfter];\n      };\n      let attached = elements.flatMap(visit);\n      const orphaned = pseudoElements.filter((pseudo) => !placed.has(pseudo));\n      if (orphaned.length > 0) {\n        const byParent = /* @__PURE__ */ new Map();\n        for (const pseudo of orphaned) {\n          if (!pseudo.source?.parentId) continue;\n          const bucket = byParent.get(pseudo.source.parentId) ?? [];\n          bucket.push(pseudo);\n          byParent.set(pseudo.source.parentId, bucket);\n        }\n        const compareChildPaths = (left, right) => {\n          const leftPath = left.source?.domPath ?? "";\n          const rightPath = right.source?.domPath ?? "";\n          return leftPath.localeCompare(rightPath, void 0, { numeric: true });\n        };\n        const insert = (element) => {\n          if (element.type !== "container") return element;\n          let children = element.children.map(insert);\n          const additions = element.source?.id ? byParent.get(element.source.id) : void 0;\n          if (additions) {\n            additions.forEach((pseudo) => placed.add(pseudo));\n            children = [...children, ...additions].sort(compareChildPaths);\n          }\n          return { ...element, children };\n        };\n        attached = attached.map(insert);\n      }\n      return [\n        ...attached,\n        ...pseudoElements.filter((pseudo) => !placed.has(pseudo))\n      ];\n    }\n    const allSections = Array.from(root.querySelectorAll("section")).filter(\n      (section) => {\n        if (section.parentElement?.closest("section")) return false;\n        if (section.parentElement?.tagName.toLowerCase() === "foreignobject") {\n          return true;\n        }\n        return section.hasAttribute("data-marpit-pagination");\n      }\n    );\n    const allSvgs = Array.from(\n      document.querySelectorAll("svg[data-marpit-svg]")\n    );\n    const slideGroups = /* @__PURE__ */ new Map();\n    for (const [index, section] of allSections.entries()) {\n      const fo = section.parentElement;\n      const svg = fo?.tagName.toLowerCase() === "foreignobject" ? fo.parentElement : null;\n      const key = svg?.hasAttribute("data-marpit-svg") ? String(allSvgs.indexOf(svg)) : section.getAttribute("data-marpit-pagination") ?? section.getAttribute("id") ?? String(index);\n      const layer = section.getAttribute("data-marpit-advanced-background");\n      if (!slideGroups.has(key)) slideGroups.set(key, {});\n      const entry = slideGroups.get(key);\n      if (layer === "content") {\n        entry.content = section;\n      } else if (layer === "background") {\n        entry.background = section;\n      } else if (layer === "pseudo") {\n        entry.pseudo = section;\n      } else {\n        entry.content = section;\n      }\n    }\n    for (const { content } of slideGroups.values()) {\n      const sec = content;\n      if (!sec) continue;\n      const secClass = sec.className?.trim() ?? "";\n      if (secClass) continue;\n      for (const pseudo of ["::before", "::after"]) {\n        const ps = getComputedStyle(sec, pseudo);\n        const rawC = ps.content;\n        if (!rawC || rawC === "none" || rawC === "normal") continue;\n        if (rawC.replace(/^["\']|["\']$/g, "").trim() !== "") continue;\n        const bg = ps.backgroundColor;\n        if (!bg || bg === "transparent" || bg === "rgba(0, 0, 0, 0)") continue;\n        globalPseudoSignatures.add(`${pseudo}:${bg}`);\n      }\n    }\n    return Array.from(slideGroups.values()).map(\n      ({ content, background, pseudo }, slideIdx) => {\n        const section = content ?? background;\n        sourceRefFor(section);\n        const sectionStyle = getComputedStyle(section);\n        let slideRect;\n        if (background && content) {\n          const fo = section.parentElement;\n          const svg = fo?.tagName.toLowerCase() === "foreignobject" ? fo.parentElement : null;\n          slideRect = svg?.hasAttribute("data-marpit-svg") ? svg.getBoundingClientRect() : section.getBoundingClientRect();\n        } else {\n          slideRect = section.getBoundingClientRect();\n        }\n        const sectionRect = slideRect;\n        const backgroundImages = [];\n        if (background) {\n          const figures = background.querySelectorAll("figure");\n          for (const fig of Array.from(figures)) {\n            const figStyle = getComputedStyle(fig);\n            if (!figStyle.backgroundImage || figStyle.backgroundImage === "none")\n              continue;\n            const urlMatch = figStyle.backgroundImage.match(\n              /url\\(["\']?([^"\')]+)["\']?\\)/\n            );\n            if (!urlMatch) continue;\n            const figRect = fig.getBoundingClientRect();\n            const cssFilter = figStyle.filter && figStyle.filter !== "none" ? figStyle.filter : void 0;\n            const bgSize = figStyle.backgroundSize;\n            const isContain = bgSize === "contain";\n            backgroundImages.push({\n              url: urlMatch[1],\n              x: figRect.left - sectionRect.left,\n              y: figRect.top - sectionRect.top,\n              width: figRect.width || sectionRect.width,\n              height: figRect.height || sectionRect.height,\n              ...cssFilter ? { cssFilter } : {},\n              ...isContain ? { backgroundSizeContain: true } : {},\n              pageX: figRect.left,\n              pageY: figRect.top\n            });\n          }\n        }\n        if (backgroundImages.length === 0) {\n          const bgImg = sectionStyle.backgroundImage;\n          if (bgImg && bgImg !== "none") {\n            const urlMatch = bgImg.match(/url\\(["\']?([^"\')]+)["\']?\\)/);\n            if (urlMatch) {\n              backgroundImages.push({\n                url: urlMatch[1],\n                x: 0,\n                y: 0,\n                width: sectionRect.width,\n                height: sectionRect.height,\n                pageX: sectionRect.left,\n                pageY: sectionRect.top,\n                fromCssFallback: true\n              });\n            } else if (/gradient\\s*\\(/.test(bgImg)) {\n              backgroundImages.push({\n                url: "",\n                // placeholder \u2014 will be replaced by rasterized data URL\n                x: 0,\n                y: 0,\n                width: sectionRect.width,\n                height: sectionRect.height,\n                pageX: sectionRect.left,\n                pageY: sectionRect.top,\n                fromCssFallback: true\n              });\n            }\n          }\n        }\n        return {\n          width: sectionRect.width,\n          height: sectionRect.height,\n          background: findBackgroundColor(section),\n          backgroundImages,\n          sourceHasPagination: [content, background, pseudo].some(\n            (node) => node?.hasAttribute("data-marpit-pagination")\n          ),\n          elements: attachPseudoElements(\n            [\n              // Pseudo-element bars on the section itself retain slide z-order.\n              ...extractPseudoElements(section, sectionRect),\n              ...walkSectionNodes(section, sectionRect)\n            ],\n            extractDescendantPseudoElements(section, sectionRect)\n          ),\n          notes: (section.querySelector("[data-marpit-presenter-notes]") ?? root.querySelector(`.bespoke-marp-note[data-index="${slideIdx}"]`))?.textContent?.trim() ?? ""\n        };\n      }\n    );\n  }\n  return __toCommonJS(dom_walker_exports);\n})();\n\nglobalThis.extractSlides = DomWalker.extractSlides;\n';

// node_modules/marp-to-editable-pptx/src/native-pptx/slide-builder.ts
var import_node_url = require("node:url");
var import_pptxgenjs = __toESM(require("pptxgenjs"));

// node_modules/marp-to-editable-pptx/src/native-pptx/renderer-profiles/windows-powerpoint.json
var windows_powerpoint_default = {
  id: "windows-powerpoint",
  text: {
    cjkFontSizeScale: 0.98,
    singleLineLeadingDivisor: 1.15,
    multiLineLeadingDivisor: 1.2,
    minimumLineSpacingMultiple: 0.85,
    fontOverrides: [
      {
        fonts: ["Segoe UI", "Yu Gothic UI", "Meiryo", "Arial"],
        glyphCategory: "cjk",
        fontSizeScale: 0.98
      }
    ]
  },
  layout: {
    columnWidthScale: 1.05,
    nowrapWidthScale: 1.15,
    listCjkFontSizeScale: 0.94,
    badgeAverageGlyphWidthEm: 0.58,
    defaultBulletIndentPt: 18
  },
  code: {
    lineSpacing: 1.2
  }
};

// node_modules/marp-to-editable-pptx/src/native-pptx/renderer-profile.ts
var WINDOWS_POWERPOINT_PROFILE = windows_powerpoint_default;
function resolveRendererProfile(renderer = "windows-powerpoint") {
  if (renderer === WINDOWS_POWERPOINT_PROFILE.id) {
    return WINDOWS_POWERPOINT_PROFILE;
  }
  throw new Error(`Unsupported renderer profile: ${renderer}`);
}
function resolveFontSizeScale(profile, fontFamily, glyphCategory) {
  const normalizedFamily = fontFamily?.toLowerCase() ?? "";
  const override = profile.text.fontOverrides?.find(
    (candidate) => candidate.glyphCategory === glyphCategory && candidate.fonts.some(
      (font) => normalizedFamily.includes(font.toLowerCase())
    )
  );
  if (override) return override.fontSizeScale;
  return glyphCategory === "cjk" ? profile.text.cjkFontSizeScale : 1;
}

// node_modules/marp-to-editable-pptx/src/native-pptx/utils.ts
function rgbToHex(rgb) {
  if (!rgb) return "000000";
  const match = rgb.match(
    /rgba?\(\s*(-?\d+)\s*,\s*(-?\d+)\s*,\s*(-?\d+)\s*(?:,\s*[\d.]+\s*)?\)/
  );
  if (!match) return "000000";
  const r = parseInt(match[1], 10);
  const g = parseInt(match[2], 10);
  const b = parseInt(match[3], 10);
  return [r, g, b].map((c) => Math.max(0, Math.min(255, c)).toString(16).padStart(2, "0")).join("").toUpperCase();
}
var genericFontMap = {
  "sans-serif": "Calibri",
  serif: "Cambria",
  monospace: "Courier New",
  cursive: "Calibri",
  fantasy: "Calibri",
  "ui-sans-serif": "Calibri",
  "ui-serif": "Cambria",
  "ui-monospace": "Courier New"
};
var systemAliases = /* @__PURE__ */ new Set([
  "-apple-system",
  "blinkmacsystemfont",
  "system-ui",
  "ui-rounded",
  "ui-monospace"
]);
var macOnlyFonts = /* @__PURE__ */ new Set([
  "sfmono-regular",
  "sf mono",
  "menlo",
  "sf pro text",
  "sf pro display"
]);
var knownSystemFonts = {
  "segoe ui": "Segoe UI",
  "helvetica neue": "Arial",
  helvetica: "Arial"
};
var japaneseFontPattern = /(noto sans jp|noto sans cjk jp|yu gothic ui|yu gothic|meiryo|biz udpgothic|biz udgothic|ms pgothic|ms gothic|hiragino sans)/i;
var japaneseTextPattern = /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff01-\uff9f]/;
var proprietaryFontPattern = /\b\d{2,4}[A-Z]{2,}/i;
function cleanFontFamily(css, sampleText) {
  if (!css) return "Calibri";
  const families = css.split(",");
  let fallback;
  const candidates = [];
  const hasJapaneseText = typeof sampleText === "string" && japaneseTextPattern.test(sampleText);
  for (const raw of families) {
    const name = raw.trim().replace(/^["']|["']$/g, "");
    if (!name) continue;
    const lower = name.toLowerCase();
    if (systemAliases.has(lower)) continue;
    if (macOnlyFonts.has(lower)) continue;
    const known = knownSystemFonts[lower];
    if (known) {
      candidates.push(known);
      continue;
    }
    const generic = genericFontMap[lower];
    if (generic) {
      if (!fallback) fallback = generic;
      continue;
    }
    candidates.push(name);
  }
  if (hasJapaneseText) {
    const japaneseCandidate = candidates.find(
      (candidate) => japaneseFontPattern.test(candidate)
    );
    if (japaneseCandidate) return japaneseCandidate;
    if (candidates.includes("Segoe UI")) return "Yu Gothic UI";
    const nonProprietary = candidates.find(
      (c) => !proprietaryFontPattern.test(c)
    );
    if (nonProprietary) return nonProprietary;
    return "Meiryo";
  }
  if (candidates.length > 0) {
    const safeFront = candidates.find((c) => !proprietaryFontPattern.test(c));
    if (safeFront) return safeFront;
    return candidates[0];
  }
  return fallback ?? "Calibri";
}
function pxToInches(px) {
  return px / 96;
}
function pxToPoints(px) {
  return px * 0.75;
}
function compositeOver(color, bg) {
  const m = color.match(
    /rgba\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*,\s*([\d.]+)\s*\)/
  );
  if (!m) return color;
  const a = parseFloat(m[4]);
  const bgM = bg.match(/rgb\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)/);
  const bgR = bgM ? parseInt(bgM[1], 10) : 255;
  const bgG = bgM ? parseInt(bgM[2], 10) : 255;
  const bgB = bgM ? parseInt(bgM[3], 10) : 255;
  const cr = Math.round(parseInt(m[1], 10) * a + bgR * (1 - a));
  const cg = Math.round(parseInt(m[2], 10) * a + bgG * (1 - a));
  const cb = Math.round(parseInt(m[3], 10) * a + bgB * (1 - a));
  return `rgb(${cr}, ${cg}, ${cb})`;
}
function isTransparent(color) {
  if (!color) return true;
  if (color === "transparent") return true;
  const match = color.match(
    /rgba\(\s*\d+\s*,\s*\d+\s*,\s*\d+\s*,\s*([\d.]+)\s*\)/
  );
  if (match) return parseFloat(match[1]) <= 0.01;
  return false;
}
function sanitizeText(text) {
  return text.replace(
    // eslint-disable-next-line no-control-regex
    /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x9F\uFEFF\u200B\u200C\u2060\uFE0E\uFE0F]/g,
    ""
  );
}

// node_modules/marp-to-editable-pptx/src/native-pptx/render-plan-helpers.ts
var EMBEDDABLE_IN_SHAPE = /* @__PURE__ */ new Set([
  "paragraph",
  "heading",
  "list",
  "blockquote"
]);
var GROUPABLE_TYPES = /* @__PURE__ */ new Set(["paragraph", "blockquote"]);
var MAX_GROUP_GAP_PX = 64;
var CONTAINMENT_SLOP_PX = 4;
var MAX_X_DIFF_PX = 5;
var MAX_WIDTH_DIFF_PX = 5;
function isSimpleTextContainer(element) {
  if (!element.children || element.children.length === 0) return false;
  if (element.runs && element.runs.length > 0) return false;
  if (element.style.borderBottom) return false;
  if (element.children.some(
    (child) => child.type === "paragraph" && child.measuredLines !== void 0 && child.measuredLines.length > 1
  ))
    return false;
  if (element.children.some(
    (child) => child.type === "list" && child.availableWidth !== void 0 && child.availableWidth > child.width + 1
  ))
    return false;
  if (!element.children.every((child) => EMBEDDABLE_IN_SHAPE.has(child.type)))
    return false;
  if (element.children.length > 1) {
    const xs = element.children.map((child) => child.x);
    if (Math.max(...xs) - Math.min(...xs) > 50) return false;
  }
  return true;
}
function isContainedIn(text, container) {
  const strict = text.x >= container.x - CONTAINMENT_SLOP_PX && text.y >= container.y - CONTAINMENT_SLOP_PX && text.x + text.width <= container.x + container.width + CONTAINMENT_SLOP_PX && text.y + text.height <= container.y + container.height + CONTAINMENT_SLOP_PX;
  if (strict) return true;
  const dx = Math.abs(text.x - container.x);
  const dy = Math.abs(text.y - container.y);
  const dw = Math.abs(text.width - container.width);
  const dh = Math.abs(text.height - container.height);
  return dx <= 16 && dy <= 16 && dw <= 16 && dh <= 16;
}
function sourceOwns(text, container) {
  if (!text.source || !container.source) return false;
  return text.source.id === container.source.id || text.source.parentId === container.source.id || text.source.domPath.startsWith(`${container.source.domPath}/`);
}
function hasKnownDifferentOwner(text, container) {
  return !!(text.source && container.source) && !sourceOwns(text, container);
}
function isEmbeddableContainer(element, slideW, slideH) {
  if (element.runs && element.runs.length > 0) return false;
  if (element.style.borderBottom) return false;
  if (element.children.length > 0 && !isSimpleTextContainer(element)) {
    return false;
  }
  const hasBackground = !isTransparent(element.style.backgroundColor);
  const hasBorderLeft = !!element.style.borderLeft && element.style.borderLeft.width > 0;
  const hasBorder = Object.keys(element.style.borders ?? {}).length > 0 || (element.style.borderWidth ?? 0) > 0 && !!element.style.borderColor && !isTransparent(element.style.borderColor);
  const hasShadow = element.style.boxShadow === true;
  if (!hasBackground && !hasBorder && !hasShadow && !hasBorderLeft) return false;
  return !(slideW > 0 && slideH > 0 && element.width >= slideW * 0.9 && element.height >= slideH * 0.9);
}
function associateContainerText(elements, slideW, slideH) {
  const result = /* @__PURE__ */ new Map();
  const containers = elements.filter(
    (element) => element.type === "container" && isEmbeddableContainer(element, slideW, slideH)
  ).sort((left, right) => left.width * left.height - right.width * right.height);
  const textElements = elements.filter(
    (element) => EMBEDDABLE_IN_SHAPE.has(element.type) && !(element.type === "paragraph" && element.measuredLines !== void 0 && element.measuredLines.length > 1)
  );
  const assigned = /* @__PURE__ */ new Set();
  for (const container of containers) {
    const owned = textElements.filter(
      (element) => !assigned.has(element) && sourceOwns(element, container)
    ).sort((left, right) => left.y - right.y);
    if (owned.length > 0) {
      result.set(container, owned);
      owned.forEach((element) => assigned.add(element));
    }
  }
  for (const container of containers) {
    if (result.has(container)) continue;
    const inside = textElements.filter(
      (element) => !assigned.has(element) && !hasKnownDifferentOwner(element, container) && isContainedIn(element, container)
    ).sort((left, right) => left.y - right.y);
    if (inside.length > 0) {
      result.set(container, inside);
      inside.forEach((element) => assigned.add(element));
    }
  }
  return result;
}
function sourceOwnerKey(element) {
  if (!element.source) return void 0;
  return element.source.parentId ?? element.source.domPath;
}
function canMergeElements(left, right) {
  if (!GROUPABLE_TYPES.has(left.type) || !GROUPABLE_TYPES.has(right.type))
    return false;
  const leftOwner = sourceOwnerKey(left);
  const rightOwner = sourceOwnerKey(right);
  if (leftOwner && rightOwner && leftOwner !== rightOwner) return false;
  for (const element of [left, right]) {
    if (element.type === "blockquote" && element.borderLeft && element.borderLeft.width > 0)
      return false;
  }
  for (const element of [left, right]) {
    if (element.type !== "paragraph") continue;
    if (element.measuredLines && element.measuredLines.length > 1) return false;
    if (element.style.lineHeight > 0 && element.height > element.style.lineHeight * 1.5)
      return false;
  }
  if (Math.abs(left.x - right.x) > MAX_X_DIFF_PX) return false;
  if (Math.abs(left.width - right.width) > MAX_WIDTH_DIFF_PX) return false;
  const gap = right.y - (left.y + left.height);
  if (gap < 0 || gap > MAX_GROUP_GAP_PX) return false;
  return true;
}
function groupAdjacentTextElements(elements) {
  if (elements.length === 0) return [];
  const groups = [[elements[0]]];
  for (const current of elements.slice(1)) {
    const lastGroup = groups[groups.length - 1];
    const previous = lastGroup[lastGroup.length - 1];
    if (canMergeElements(previous, current)) lastGroup.push(current);
    else groups.push([current]);
  }
  return groups;
}

// node_modules/marp-to-editable-pptx/src/native-pptx/pptx-plan.ts
var TEXT_ELEMENT_TYPES = /* @__PURE__ */ new Set([
  "heading",
  "paragraph",
  "list",
  "table",
  "code",
  "blockquote",
  "header",
  "footer"
]);
function assertFiniteGeometry(element) {
  for (const field of ["x", "y", "width", "height"]) {
    if (!Number.isFinite(element[field])) {
      throw new Error(`invalid ${field} geometry for ${element.type}`);
    }
  }
  if (element.type === "container") {
    element.children.forEach(assertFiniteGeometry);
  }
}
function visibleElement(element, slideHeight) {
  if (TEXT_ELEMENT_TYPES.has(element.type) && slideHeight - element.y < 20) {
    return void 0;
  }
  if (element.type !== "container") return element;
  const children = element.children.map((child) => visibleElement(child, slideHeight)).filter((child) => !!child);
  if (children.length === element.children.length && children.every((child, index) => child === element.children[index])) {
    return element;
  }
  return { ...element, children };
}
function hasVisibleContainerShape(element) {
  if (element.type !== "container") return true;
  const style = element.style;
  return !isTransparent(style.backgroundColor) || Object.keys(style.borders ?? {}).length > 0 || (style.borderWidth ?? 0) > 0 || style.borderLeft !== void 0 || style.borderBottom !== void 0 || style.boxShadow === true;
}
function applyStackingOrder(elements) {
  const beneath = [];
  const ordinary = [];
  const above = [];
  const lifted = (element) => {
    if (element.type !== "container") return element;
    const children = [];
    let changed = false;
    for (const child of element.children) {
      const normalized = lifted(child);
      if (normalized.stackingPositioned && normalized.stackingIndex !== void 0 && normalized.stackingIndex !== 0) {
        changed = true;
        if (normalized.stackingIndex < 0) beneath.push(normalized);
        else above.push(normalized);
        continue;
      }
      if (normalized !== child) changed = true;
      children.push(normalized);
    }
    return changed ? { ...element, children } : element;
  };
  for (const element of elements) {
    const normalized = lifted(element);
    if (normalized.stackingPositioned && normalized.stackingIndex !== void 0 && normalized.stackingIndex !== 0) {
      if (normalized.stackingIndex < 0) beneath.push(normalized);
      else above.push(normalized);
    } else {
      ordinary.push(normalized);
    }
  }
  beneath.sort(
    (left, right) => left.stackingIndex - right.stackingIndex
  );
  above.sort((left, right) => left.stackingIndex - right.stackingIndex);
  return [...beneath, ...ordinary, ...above];
}
function classifyElements(elements) {
  const result = { nativeCount: 0, rasterCount: 0, nativeArea: 0, totalArea: 0 };
  for (const element of elements) {
    const ownArea = Math.max(0, element.width) * Math.max(0, element.height);
    const isRaster = element.type === "image" && element.rasterize === true;
    const ownsNativeShape = !isRaster && hasVisibleContainerShape(element);
    if (isRaster) result.rasterCount++;
    else if (ownsNativeShape) {
      result.nativeCount++;
      result.nativeArea += ownArea;
    }
    if (ownsNativeShape || isRaster) result.totalArea += ownArea;
    if (element.type === "container") {
      const children = classifyElements(element.children);
      result.nativeCount += children.nativeCount;
      result.rasterCount += children.rasterCount;
      result.nativeArea += children.nativeArea;
      result.totalArea += children.totalArea;
    }
  }
  return result;
}
function collectContainerAssociations(elements, slideW, slideH) {
  const result = Array.from(
    associateContainerText(elements, slideW, slideH),
    ([container, elements2]) => ({ container, elements: elements2 })
  );
  for (const element of elements) {
    if (element.type === "container") {
      result.push(
        ...collectContainerAssociations(element.children, slideW, slideH)
      );
    }
  }
  return result;
}
function createPptxRenderPlan(slides) {
  const plannedSlides = slides.map((slide) => {
    slide.elements.forEach(assertFiniteGeometry);
    const backgrounds = slide.backgroundImages ?? [];
    const firstBackground = backgrounds[0];
    const isFullSlide = firstBackground && !firstBackground.cssFilter && !firstBackground.backgroundSizeContain && firstBackground.x <= 1 && firstBackground.y <= 1 && Math.abs(firstBackground.width - slide.width) <= 2 && Math.abs(firstBackground.height - slide.height) <= 2;
    const visibleElements = applyStackingOrder(
      slide.elements.map((element) => visibleElement(element, slide.height)).filter((element) => !!element)
    );
    const classification = classifyElements(visibleElements);
    const associations = collectContainerAssociations(
      visibleElements,
      slide.width,
      slide.height
    );
    const cssIsFallbackWhite = isTransparent(slide.background) || rgbToHex(slide.background).toUpperCase() === "FFFFFF";
    const visualBgMayBeDark = backgrounds.some(
      (background) => background.url !== "" && !background.fromCssFallback && !background.backgroundSizeContain && background.width >= slide.width * 0.8
    ) && cssIsFallbackWhite;
    return {
      source: slide,
      background: isFullSlide && backgrounds.length === 1 ? { mode: "native-slide", sourceUrl: firstBackground.url } : { mode: "shapes" },
      visibleElements,
      elementGroups: groupAdjacentTextElements(visibleElements),
      containerAssociations: associations,
      visualBgMayBeDark,
      editableElementCount: classification.nativeCount,
      rasterizedElementCount: classification.rasterCount,
      editableAreaRatio: classification.totalArea > 0 ? Math.min(1, classification.nativeArea / classification.totalArea) : 1
    };
  });
  const editableElementCount = plannedSlides.reduce(
    (sum, slide) => sum + slide.editableElementCount,
    0
  );
  const rasterizedElementCount = plannedSlides.reduce(
    (sum, slide) => sum + slide.rasterizedElementCount,
    0
  );
  const totals = plannedSlides.reduce(
    (aggregate, slide) => {
      const classification = classifyElements(slide.visibleElements);
      aggregate.editableArea += classification.nativeArea;
      aggregate.totalArea += classification.totalArea;
      return aggregate;
    },
    { editableArea: 0, totalArea: 0 }
  );
  return {
    slides: plannedSlides,
    totals: {
      editableElementCount,
      rasterizedElementCount,
      editableAreaRatio: totals.totalArea > 0 ? totals.editableArea / totals.totalArea : 1
    }
  };
}

// node_modules/marp-to-editable-pptx/src/native-pptx/slide-builder.ts
var japaneseTextPattern2 = /[\u3000-\u30ff\u3400-\u9fff\uf900-\ufaff\uff01-\uff9f]/;
var rendererProfile = resolveRendererProfile();
function cssBorderStyleToDash(borderStyle) {
  if (!borderStyle || borderStyle === "solid") return void 0;
  if (borderStyle === "dashed") return "dash";
  if (borderStyle === "dotted") return "sysDot";
  return void 0;
}
function drawPreciseContainerBorders(slide, borders, x, y, width, height) {
  const edges = [
    borders.top ? {
      border: borders.top,
      x,
      y,
      width,
      height: pxToInches(borders.top.width),
      line: { x, y, w: width, h: 0 }
    } : void 0,
    borders.right ? {
      border: borders.right,
      x: x + width - pxToInches(borders.right.width),
      y,
      width: pxToInches(borders.right.width),
      height,
      line: { x: x + width, y, w: 0, h: height }
    } : void 0,
    borders.bottom ? {
      border: borders.bottom,
      x,
      y: y + height - pxToInches(borders.bottom.width),
      width,
      height: pxToInches(borders.bottom.width),
      line: { x, y: y + height, w: width, h: 0 }
    } : void 0,
    borders.left ? {
      border: borders.left,
      x,
      y,
      width: pxToInches(borders.left.width),
      height,
      line: { x, y, w: 0, h: height }
    } : void 0
  ];
  for (const edge of edges) {
    if (!edge) continue;
    const color = rgbToHex(edge.border.color);
    const dashType = cssBorderStyleToDash(edge.border.style);
    if (dashType) {
      slide.addShape("line", {
        x: edge.line.x,
        y: edge.line.y,
        w: edge.line.w,
        h: edge.line.h,
        line: {
          color,
          width: Math.max(0.25, pxToPoints(edge.border.width)),
          dashType
        }
      });
    } else {
      slide.addShape("rect", {
        x: edge.x,
        y: edge.y,
        w: edge.width,
        h: edge.height,
        fill: { color }
      });
    }
  }
}
function resolveImageSource(url) {
  if (url.startsWith("data:")) return { data: url };
  if (url.startsWith("file:")) return { path: (0, import_node_url.fileURLToPath)(url) };
  return { path: url };
}
function computeLineSpacing(style, multiLine, profile = rendererProfile) {
  const { lineHeight, fontSize } = style;
  if (!lineHeight || !fontSize || lineHeight <= 0 || fontSize <= 0)
    return void 0;
  const m = lineHeight / fontSize;
  if (m < 0.5 || m > 4) return void 0;
  const divisor = multiLine ? profile.text.multiLineLeadingDivisor : profile.text.singleLineLeadingDivisor;
  const adjusted = Math.max(
    profile.text.minimumLineSpacingMultiple,
    m / divisor
  );
  return Math.round(adjusted * 100) / 100;
}
function scaledFontSizePx(text, fontSizePx = 16, fontFamily) {
  const glyphCategory = japaneseTextPattern2.test(text) ? "cjk" : "latin";
  const scale = resolveFontSizeScale(
    rendererProfile,
    fontFamily,
    glyphCategory
  );
  return fontSizePx * scale;
}
function computeBadgeRunFontScale(runs, widthPx, heightPx) {
  const text = runs.filter((run2) => !run2.breakLine).map((run2) => run2.text).join("").trim();
  if (!text || /[\p{Extended_Pictographic}\u200d]/u.test(text)) return 1;
  const baseFontPx = runs.find((run2) => run2.fontSize)?.fontSize ?? 16;
  const maxCharsPerLine = Math.max(
    1,
    Math.floor(
      widthPx / (baseFontPx * rendererProfile.layout.badgeAverageGlyphWidthEm)
    )
  );
  const estimatedLines = Math.max(1, Math.ceil(text.length / maxCharsPerLine));
  const maxFontPx = heightPx * 0.75 / (estimatedLines * 1);
  return Math.min(1, maxFontPx / (baseFontPx * 0.75));
}
function computeCharSpacing(style) {
  const ls = style.letterSpacing;
  if (!ls || Math.abs(ls) < 0.1) return void 0;
  return Math.round(pxToPoints(ls) * 100) / 100;
}
function computeTextInset(style) {
  const pt = (style.paddingTop ?? 0) * 0.75;
  const pr = (style.paddingRight ?? 0) * 0.75;
  const pb = (style.paddingBottom ?? 0) * 0.75;
  const pl = (style.paddingLeft ?? 0) * 0.75;
  return pt || pr || pb || pl ? [pl, pr, pb, pt] : 0;
}
function createSlideNumberProps(slideW, slideH) {
  const bottomInset = slideH * (30 / 720);
  const rightInset = slideW * (40 / 1280);
  const boxHeight = slideH * (24 / 720);
  const fontSize = slideH * (18 / 720);
  return {
    x: 0,
    y: pxToInches(Math.max(0, slideH - bottomInset - boxHeight)),
    w: pxToInches(Math.max(48, slideW - rightInset)),
    h: pxToInches(Math.max(1, boxHeight)),
    align: "right",
    color: "777777",
    fontSize: Math.round(pxToPoints(fontSize) * 100) / 100,
    margin: 0
  };
}
function cssListStyleToNumberStyle(cssType) {
  if (!cssType) return void 0;
  switch (cssType) {
    case "decimal":
      return "arabicPeriod";
    case "lower-alpha":
    case "lower-latin":
      return "alphaLcPeriod";
    case "upper-alpha":
    case "upper-latin":
      return "alphaUcPeriod";
    case "lower-roman":
      return "romanLcPeriod";
    case "upper-roman":
      return "romanUcPeriod";
    default:
      return void 0;
  }
}
function cssListStyleToBulletChar(cssType) {
  if (!cssType) return void 0;
  switch (cssType) {
    case "circle":
      return "25E6";
    // ◦ (white bullet — smaller than ○ U+25CB)
    case "square":
      return "25AA";
    // ▪ (black small square)
    case "none":
      return "200B";
    // zero-width space (invisible)
    default:
      return void 0;
  }
}
function createListBulletOption(item, ordered, continuation = false, startNumber, listStyleType) {
  const extraIndent = item.leadingOffset ? pxToPoints(item.leadingOffset) : 0;
  const indent = Math.round(
    (rendererProfile.layout.defaultBulletIndentPt + extraIndent) * 100
  ) / 100;
  const effectiveStyle = item.listStyleType ?? listStyleType;
  if (continuation) {
    return {
      characterCode: "200B",
      ...extraIndent > 0 ? { indent } : {}
    };
  }
  const isUnorderedStyle = effectiveStyle === "disc" || effectiveStyle === "circle" || effectiveStyle === "square" || effectiveStyle === "none";
  const isOrderedStyle = effectiveStyle === "decimal" || effectiveStyle === "lower-alpha" || effectiveStyle === "lower-latin" || effectiveStyle === "upper-alpha" || effectiveStyle === "upper-latin" || effectiveStyle === "lower-roman" || effectiveStyle === "upper-roman";
  const effectiveOrdered = isUnorderedStyle ? false : isOrderedStyle ? true : ordered;
  if (effectiveOrdered) {
    const numberStyle = cssListStyleToNumberStyle(effectiveStyle) ?? "arabicPeriod";
    return {
      type: "number",
      style: numberStyle,
      ...startNumber !== void 0 ? { numberStartAt: startNumber } : {},
      indent
    };
  }
  const bulletChar = cssListStyleToBulletChar(effectiveStyle);
  if (bulletChar) {
    return {
      characterCode: bulletChar,
      ...extraIndent > 0 ? { indent } : {}
    };
  }
  return extraIndent > 0 ? { indent } : true;
}
function executePptxRenderPlan(renderPlan) {
  const pptx = new import_pptxgenjs.default();
  const slides = renderPlan.slides.map((slide) => slide.source);
  const slideW = slides[0]?.width ?? 1280;
  const slideH = slides[0]?.height ?? 720;
  pptx.defineLayout({
    name: "MARP",
    width: pxToInches(slideW),
    height: pxToInches(slideH)
  });
  pptx.layout = "MARP";
  const useSlideNumbers = slides.some((slide) => slide.sourceHasPagination);
  for (const slidePlan of renderPlan.slides) {
    const slideData = slidePlan.source;
    const slide = pptx.addSlide();
    if (useSlideNumbers) {
      slide.slideNumber = {
        ...createSlideNumberProps(slideW, slideH)
      };
    }
    const bgColor = isTransparent(slideData.background) ? "FFFFFF" : rgbToHex(slideData.background);
    const bgImages = slideData.backgroundImages ?? [];
    if (slidePlan.background.mode === "native-slide") {
      slide.background = resolveImageSource(slidePlan.background.sourceUrl);
    } else {
      slide.background = { fill: bgColor };
      for (const bg of bgImages) {
        const x = pxToInches(bg.x);
        const y = pxToInches(bg.y);
        const w = pxToInches(bg.width);
        const h = pxToInches(bg.height);
        const imgOpts = {
          x,
          y,
          w,
          h,
          ...resolveImageSource(bg.url)
        };
        slide.addImage(imgOpts);
      }
    }
    const slideBgColor = slideData.background ?? "rgb(255, 255, 255)";
    const visualBgMayBeDark = slidePlan.visualBgMayBeDark;
    const containerAssoc = new Map(
      slidePlan.containerAssociations.map(({ container, elements }) => [
        container,
        elements
      ])
    );
    const embeddedByContainer = new Set(
      Array.from(containerAssoc.values()).flat()
    );
    const visibleElements = slidePlan.visibleElements;
    const groups = slidePlan.elementGroups;
    for (const group of groups) {
      const active = group.filter((el) => !embeddedByContainer.has(el));
      if (active.length === 0) continue;
      if (active.length === 1) {
        placeElement(
          slide,
          active[0],
          slideData.width,
          slideData.height,
          slideBgColor,
          visualBgMayBeDark,
          containerAssoc
        );
      } else {
        placeGroupedTextElements(
          slide,
          active,
          slideData.width,
          slideData.height,
          slideBgColor,
          visualBgMayBeDark
        );
      }
    }
    if (slideData.notes) {
      slide.addNotes(slideData.notes);
    }
  }
  return pptx;
}
var TEXT_ELEMENT_TYPES2 = /* @__PURE__ */ new Set([
  "heading",
  "paragraph",
  "list",
  "blockquote",
  "code",
  "table",
  "header",
  "footer"
]);
var EMBEDDABLE_IN_SHAPE2 = /* @__PURE__ */ new Set([
  "paragraph",
  "heading",
  "list",
  "blockquote"
]);
function isSimpleTextContainer2(el) {
  if (!el.children || el.children.length === 0) return false;
  if (el.runs && el.runs.length > 0) return false;
  if (el.style.borderBottom) return false;
  if (el.children.some(
    (child) => child.type === "paragraph" && child.measuredLines !== void 0 && child.measuredLines.length > 1
  ))
    return false;
  if (el.children.some(
    (child) => child.type === "list" && child.availableWidth !== void 0 && child.availableWidth > child.width + 1
  ))
    return false;
  if (!el.children.every((c) => EMBEDDABLE_IN_SHAPE2.has(c.type))) return false;
  if (el.children.length > 1) {
    const xs = el.children.map((c) => c.x);
    const xSpan = Math.max(...xs) - Math.min(...xs);
    if (xSpan > 50) return false;
  }
  return true;
}
function computeContainerInset(el) {
  const { children } = el;
  if (!children || children.length === 0) return 0;
  return computeInsetFromElements(el, children);
}
function buildContainerEmbeddedRuns(el, containerBg, slideBg, visualBgMayBeDark) {
  const bgHex = containerBg ? rgbToHex(containerBg) : void 0;
  const toTP = (run2) => toTextProps(
    bgHex && run2.backgroundColor && rgbToHex(run2.backgroundColor) === bgHex ? { ...run2, backgroundColor: void 0 } : run2,
    slideBg,
    visualBgMayBeDark
  );
  const allRuns = [];
  for (let i = 0; i < el.children.length; i++) {
    const child = el.children[i];
    if (i > 0) allRuns.push({ text: "", options: { breakLine: true } });
    if (child.type === "list") {
      const listEl = child;
      let topLevelCount = 0;
      const listRuns = listEl.items.flatMap((item, idx) => {
        const startNum = item.level === 0 && listEl.ordered ? (listEl.startNumber ?? 1) + topLevelCount : void 0;
        if (item.level === 0) topLevelCount++;
        return toListTextProps(
          item,
          listEl.ordered,
          idx < listEl.items.length - 1,
          slideBg,
          visualBgMayBeDark,
          startNum,
          listEl.listStyleType
        );
      });
      allRuns.push(...listRuns);
    } else if ("runs" in child && Array.isArray(child.runs)) {
      allRuns.push(...child.runs.map(toTP));
    }
  }
  return allRuns;
}
function associateContainerText2(elements, slideW, slideH) {
  return associateContainerText(elements, slideW, slideH);
}
function buildRunsFromElements(elements, containerBg, slideBg, visualBgMayBeDark) {
  const toTP = (r) => toTextProps(r, slideBg, visualBgMayBeDark);
  const bgHex = containerBg ? rgbToHex(containerBg) : void 0;
  const allRuns = [];
  for (let i = 0; i < elements.length; i++) {
    const el = elements[i];
    if (i > 0) allRuns.push({ text: "", options: { breakLine: true } });
    if (el.type === "list") {
      const listEl = el;
      let topLevelCount = 0;
      const listRuns = listEl.items.flatMap((item, idx) => {
        const startNum = item.level === 0 && listEl.ordered ? (listEl.startNumber ?? 1) + topLevelCount : void 0;
        if (item.level === 0) topLevelCount++;
        return toListTextProps(
          item,
          listEl.ordered,
          idx < listEl.items.length - 1,
          slideBg,
          visualBgMayBeDark,
          startNum,
          listEl.listStyleType
        );
      });
      allRuns.push(...listRuns);
    } else if ("runs" in el && Array.isArray(el.runs)) {
      const runs = el.runs;
      const mapped = runs.map((r) => {
        if (bgHex && !r.breakLine && r.backgroundColor && rgbToHex(r.backgroundColor) === bgHex) {
          return { ...r, backgroundColor: void 0 };
        }
        return r;
      });
      allRuns.push(...mapped.map(toTP));
    }
  }
  return allRuns;
}
function computeInsetFromElements(container, elements) {
  if (elements.length === 0) return 0;
  const firstEl = elements[0];
  const lastEl = elements[elements.length - 1];
  const padding = (el) => {
    const style = "style" in el ? el.style : void 0;
    return {
      left: style?.paddingLeft ?? 0,
      right: style?.paddingRight ?? 0,
      bottom: style?.paddingBottom ?? 0,
      top: style?.paddingTop ?? 0
    };
  };
  const firstPadding = padding(firstEl);
  const lastPadding = padding(lastEl);
  const outerWidth = (element) => element.type === "list" ? element.availableWidth ?? element.width : element.width;
  const topPx = Math.max(0, firstEl.y - container.y + firstPadding.top);
  const leftPx = Math.max(
    0,
    Math.min(...elements.map((e) => e.x - container.x + padding(e).left))
  );
  const rightPx = Math.max(
    0,
    Math.min(
      ...elements.map(
        (e) => container.x + container.width - (e.x + outerWidth(e)) + padding(e).right
      )
    )
  );
  const bottomPx = Math.max(
    0,
    container.y + container.height - (lastEl.y + lastEl.height - lastPadding.bottom)
  );
  const top = topPx * 0.75;
  const left = leftPx * 0.75;
  const right = rightPx * 0.75;
  const bottom = bottomPx * 0.75;
  return top || left || right || bottom ? [left, right, bottom, top] : 0;
}
function computeHighlight(backgroundColor, textColor, slideBg, visualBgMayBeDark = false) {
  if (!backgroundColor) return void 0;
  const composited = compositeOver(backgroundColor, slideBg);
  const hex = rgbToHex(composited);
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const bghex = rgbToHex(slideBg);
  const br = parseInt(bghex.slice(0, 2), 16);
  const bgG = parseInt(bghex.slice(2, 4), 16);
  const bgB = parseInt(bghex.slice(4, 6), 16);
  if (Math.max(Math.abs(r - br), Math.abs(g - bgG), Math.abs(b - bgB)) < 10)
    return void 0;
  if (visualBgMayBeDark && r > 200 && g > 200 && b > 200) return void 0;
  if (r > 200 && g > 200 && b > 200 && textColor) {
    const tc = rgbToHex(textColor);
    const tr = parseInt(tc.slice(0, 2), 16);
    const tg = parseInt(tc.slice(2, 4), 16);
    const tb = parseInt(tc.slice(4, 6), 16);
    if (tr > 200 && tg > 200 && tb > 200) return void 0;
  }
  return hex;
}
function placeElement(slide, el, slideW = 0, slideH = 0, slideBg = "rgb(255, 255, 255)", visualBgMayBeDark = false, containerAssoc) {
  if (slideH > 0 && TEXT_ELEMENT_TYPES2.has(el.type) && slideH - el.y < 20) return;
  const x = pxToInches(el.x);
  const y = pxToInches(el.y);
  const w = pxToInches(el.width);
  const rawH = pxToInches(el.height);
  const textVerticalSlack = TEXT_ELEMENT_TYPES2.has(el.type) && "style" in el && "lineHeight" in el.style && el.height <= 30 && el.height <= el.style.lineHeight * 1.5 ? pxToInches(3) : 0;
  const h = slideH > 0 && TEXT_ELEMENT_TYPES2.has(el.type) ? Math.min(
    rawH + textVerticalSlack,
    Math.max(0.01, pxToInches(slideH) - y)
  ) : rawH + textVerticalSlack;
  const toTP = (r) => toTextProps(r, slideBg, visualBgMayBeDark);
  const elLineHeight = "style" in el && "lineHeight" in el.style ? el.style.lineHeight : 0;
  const isMultiLine = elLineHeight > 0 && el.height > elLineHeight * 1.5;
  switch (el.type) {
    case "heading": {
      const headingBorderW = el.borderLeft && el.borderLeft.width > 0 ? pxToInches(el.borderLeft.width) : 0;
      if (headingBorderW > 0) {
        slide.addShape("rect", {
          x,
          y,
          w: headingBorderW,
          h,
          fill: { color: rgbToHex(el.borderLeft.color) },
          line: { color: rgbToHex(el.borderLeft.color) }
        });
      }
      const isFullWidthHeading = slideW > 0 && el.x < slideW * 0.15 && el.x + el.width > slideW * 0.85;
      const headingTextW = isFullWidthHeading ? Math.max(0.01, pxToInches(slideW - el.x - 16) - headingBorderW) : slideW > 0 ? Math.max(
        0.01,
        pxToInches(
          Math.min(
            el.width * rendererProfile.layout.columnWidthScale,
            slideW - el.x
          )
        ) - headingBorderW
      ) : Math.max(0.01, w - headingBorderW);
      const headingInset = computeTextInset(el.style);
      const headingHalfLeading = el.style.lineHeight > 0 && el.style.fontSize > 0 && el.style.fontSize <= 32 ? (el.style.lineHeight - el.style.fontSize) / 2 : 0;
      const headingY = headingHalfLeading > 0 ? pxToInches(el.y - headingHalfLeading) : y;
      slide.addText(
        el.runs.map((run2) => {
          const textProps = toTP(run2);
          if (el.style.fontWeight >= 500 && el.style.fontWeight < 700) {
            const fontFace = cleanFontFamily(run2.fontFamily, run2.text);
            if (fontFace === "Segoe UI") {
              textProps.options = {
                ...textProps.options,
                fontFace: "Segoe UI Semibold",
                bold: false
              };
            }
          }
          return textProps;
        }),
        {
          x: x + headingBorderW,
          y: headingY,
          w: headingTextW,
          h,
          margin: headingInset,
          valign: "top",
          align: el.style.textAlign,
          lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
          charSpacing: computeCharSpacing(el.style)
        }
      );
      if (el.borderBottom && el.borderBottom.width > 0) {
        const bh = pxToInches(el.borderBottom.width);
        const bbColor = rgbToHex(el.borderBottom.color);
        const bbDash = cssBorderStyleToDash(el.borderBottom.style);
        slide.addShape("rect", {
          x,
          y: y + h,
          w,
          h: bh,
          ...bbDash ? {
            line: {
              color: bbColor,
              width: Math.max(0.25, pxToPoints(el.borderBottom.width)),
              dashType: bbDash
            }
          } : {
            fill: { color: bbColor },
            line: { color: bbColor, width: 0.25 }
          }
        });
      }
      break;
    }
    case "paragraph": {
      if (el.measuredLines && el.measuredLines.length > 1) {
        for (const line of el.measuredLines) {
          slide.addText(
            line.runs.map(toTP),
            {
              x: pxToInches(line.x),
              y: pxToInches(line.y),
              w: pxToInches(line.width),
              h: Math.max(pxToInches(line.height), 0.01),
              margin: 0,
              valign: "top",
              align: line.textAlign ?? el.style.textAlign,
              lineSpacing: line.lineHeight ? pxToPoints(line.lineHeight) : void 0,
              paraSpaceBefore: 0,
              paraSpaceAfter: 0,
              wrap: false
            }
          );
        }
        break;
      }
      const noWrap = el.style.whiteSpace === "nowrap";
      const isSingleLine = el.style.lineHeight > 0 && el.height <= el.style.lineHeight * 1.5;
      const isNarrowGlyph = el.width <= 40 && el.style.fontSize >= el.width * 0.5;
      const glyphWidthBufferPx = isNarrowGlyph ? 8 : 0;
      let paraW;
      if (noWrap) {
        paraW = slideW > 0 ? Math.max(
          w,
          pxToInches(
            Math.min(
              el.width * rendererProfile.layout.nowrapWidthScale,
              slideW - el.x
            )
          )
        ) : w;
      } else if (slideW > 0 && isSingleLine) {
        const extendedW = el.width * rendererProfile.layout.columnWidthScale;
        const maxW = slideW - el.x - 4;
        paraW = Math.max(w, pxToInches(Math.min(extendedW, maxW)));
      } else {
        paraW = w;
      }
      if (glyphWidthBufferPx > 0 && slideW > 0) {
        paraW = pxToInches(
          Math.min(el.width + glyphWidthBufferPx, slideW - el.x)
        );
      }
      const paraX = glyphWidthBufferPx > 0 && el.style.textAlign === "center" ? pxToInches(Math.max(0, el.x - glyphWidthBufferPx / 2)) : x;
      const paraHalfLeading = el.style.lineHeight > 0 && el.style.fontSize > 0 ? (el.style.lineHeight - el.style.fontSize) / 2 : 0;
      const paraY = paraHalfLeading > 0 ? pxToInches(el.y - paraHalfLeading) : y;
      const paddingLeft = el.style.paddingLeft ?? 0;
      const paddingRight = el.style.paddingRight ?? 0;
      const paddingTop = el.style.paddingTop ?? 0;
      const paddingBottom = el.style.paddingBottom ?? 0;
      const hasBoxPadding = paddingLeft > 0 || paddingRight > 0 || paddingTop > 0 || paddingBottom > 0;
      const textX = hasBoxPadding ? pxToInches(el.x + paddingLeft) : paraX;
      const textY = hasBoxPadding ? pxToInches(el.y + paddingTop) : paraY;
      const textW = hasBoxPadding ? Math.max(0.01, pxToInches(el.width - paddingLeft - paddingRight)) : paraW;
      const textH = hasBoxPadding ? Math.max(0.01, h - pxToInches(paddingTop + paddingBottom)) : h;
      slide.addText(
        el.runs.map((run2) => {
          const props = toTP(run2);
          if (!run2.breakLine && props.options?.fontSize && el.runs.some((candidate) => candidate.breakLine) && japaneseTextPattern2.test(run2.text)) {
            props.options.fontSize *= 0.94;
          }
          return props;
        }),
        {
          x: textX,
          y: textY,
          w: textW,
          h: textH,
          margin: hasBoxPadding ? 0 : computeTextInset(el.style),
          valign: el.valign ?? "top",
          align: el.style.textAlign,
          lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
          charSpacing: computeCharSpacing(el.style),
          ...noWrap ? { wrap: false } : {}
        }
      );
      break;
    }
    case "header":
    case "footer":
      slide.addText(
        el.runs.map(toTP),
        {
          x,
          y,
          w,
          h,
          margin: 0,
          valign: "top",
          align: el.style.textAlign,
          lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
          charSpacing: computeCharSpacing(el.style)
        }
      );
      break;
    case "blockquote":
      if (el.borderLeft && el.borderLeft.width > 0) {
        const bw = pxToInches(el.borderLeft.width);
        slide.addShape("rect", {
          x,
          y,
          w: bw,
          h,
          fill: { color: rgbToHex(el.borderLeft.color) }
        });
        slide.addText(
          el.runs.map(toTP),
          {
            x: x + bw,
            y,
            w: w - bw,
            h,
            margin: computeTextInset(el.style),
            valign: "top",
            align: el.style.textAlign,
            lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
            charSpacing: computeCharSpacing(el.style)
          }
        );
      } else {
        slide.addText(
          el.runs.map(toTP),
          {
            x,
            y,
            w,
            h,
            margin: computeTextInset(el.style),
            valign: "top",
            align: el.style.textAlign,
            lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
            paraSpaceBefore: 0,
            paraSpaceAfter: 0,
            charSpacing: computeCharSpacing(el.style)
          }
        );
      }
      break;
    case "list": {
      const halfLeadingPx = el.style.lineHeight > 0 && el.style.fontSize > 0 ? (el.style.lineHeight - el.style.fontSize) / 2 : 0;
      const yShiftPx = halfLeadingPx * 1;
      const listY = pxToInches(el.y - yShiftPx);
      const listH = Math.max(0.01, h + pxToInches(yShiftPx));
      let topLevelCount = 0;
      const listW = (el.availableWidth ?? el.width) > 0 && slideW > 0 ? pxToInches(
        Math.min(
          (el.availableWidth ?? el.width) * rendererProfile.layout.columnWidthScale,
          slideW - el.x
        )
      ) : w;
      slide.addText(
        el.items.flatMap((item, index) => {
          const startNum = item.level === 0 && el.ordered ? (el.startNumber ?? 1) + topLevelCount : void 0;
          if (item.level === 0) topLevelCount++;
          return toListTextProps(
            item,
            el.ordered,
            index < el.items.length - 1,
            slideBg,
            visualBgMayBeDark,
            startNum,
            el.listStyleType
          );
        }),
        {
          x,
          y: listY,
          w: listW,
          h: listH,
          margin: 0,
          valign: "top",
          align: el.style.textAlign,
          lineSpacingMultiple: computeLineSpacing(el.style, isMultiLine),
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
          charSpacing: computeCharSpacing(el.style)
        }
      );
      break;
    }
    case "table": {
      const refCell = el.rows[0]?.cells[0];
      const hasCellPadding = refCell?.style.paddingTop !== void 0 && refCell?.style.paddingTop !== null;
      const tableFallbackMargin = hasCellPadding ? [
        pxToInches(refCell.style.paddingTop),
        pxToInches(refCell.style.paddingRight ?? 0),
        pxToInches(refCell.style.paddingBottom ?? 0),
        pxToInches(refCell.style.paddingLeft ?? 0)
      ] : [0.1, 0.05, 0.1, 0.05];
      slide.addTable(
        el.rows.map(
          (row) => row.cells.map((cell, cellIndex) => {
            const cellMargin = cell.style.paddingTop !== void 0 ? [
              pxToInches(cell.style.paddingTop ?? 0),
              pxToInches(cell.style.paddingRight ?? 0),
              pxToInches(cell.style.paddingBottom ?? 0),
              pxToInches(cell.style.paddingLeft ?? 0)
            ] : void 0;
            if (cell.runs && cell.runs.length > 0) {
              const cellOpts2 = {
                align: cell.style.textAlign,
                ...cellMargin ? { margin: cellMargin } : {},
                ...cell.colspan && cell.colspan > 1 ? { colspan: cell.colspan } : {},
                ...cell.rowspan && cell.rowspan > 1 ? { rowspan: cell.rowspan } : {}
              };
              if (!isTransparent(cell.style.backgroundColor)) {
                cellOpts2.fill = { color: rgbToHex(cell.style.backgroundColor) };
              }
              if (cell.style.borderColor && !isTransparent(cell.style.borderColor)) {
                cellOpts2.border = {
                  pt: pxToPoints(cell.style.borderWidth ?? 0) || 1,
                  color: rgbToHex(cell.style.borderColor),
                  ...cssBorderStyleToDash(cell.style.borderStyle) ? { dashType: cssBorderStyleToDash(cell.style.borderStyle) } : {}
                };
              }
              const cellEffBg = !isTransparent(cell.style.backgroundColor) ? cell.style.backgroundColor : slideBg;
              const runToTextProp = (r) => {
                const hl = computeHighlight(
                  r.backgroundColor,
                  r.color,
                  cellEffBg,
                  visualBgMayBeDark
                );
                const baseFontSizePx = scaledFontSizePx(
                  r.text,
                  r.fontSize ?? cell.style.fontSize,
                  r.fontFamily ?? cell.style.fontFamily
                );
                const columnWidth = el.colWidths?.[cellIndex] ?? el.width / row.cells.length;
                const availableWidth = Math.max(
                  1,
                  columnWidth - (cell.style.paddingLeft ?? 0) - (cell.style.paddingRight ?? 0)
                );
                const glyphWidthEm = japaneseTextPattern2.test(r.text) ? 1 : 0.58;
                const estimatedTextWidth = Array.from(r.text).length * baseFontSizePx * glyphWidthEm;
                const fitScale = Math.max(
                  0.9,
                  Math.min(1, availableWidth / Math.max(1, estimatedTextWidth))
                );
                return {
                  text: sanitizeText(r.text),
                  options: {
                    color: rgbToHex(r.color),
                    fontSize: pxToPoints(
                      baseFontSizePx * fitScale
                    ),
                    fontFace: cleanFontFamily(
                      r.fontFamily ?? cell.style.fontFamily,
                      r.text
                    ),
                    bold: r.bold ?? cell.isHeader ?? cell.style.fontWeight >= 600,
                    italic: r.italic,
                    ...hl ? { highlight: hl } : {}
                  }
                };
              };
              let textArray;
              if (cell.paragraphs && cell.paragraphs.length > 0) {
                textArray = cell.paragraphs.flatMap((para, pIdx) => {
                  const paraRuns = para.runs.filter((r) => !r.breakLine).map(runToTextProp);
                  if (paraRuns.length === 0) {
                    paraRuns.push({ text: " ", options: {
                      fontSize: pxToPoints(
                        scaledFontSizePx(
                          " ",
                          cell.style.fontSize,
                          cell.style.fontFamily
                        )
                      )
                    } });
                  }
                  if (pIdx < cell.paragraphs.length - 1) {
                    const last = paraRuns[paraRuns.length - 1];
                    paraRuns[paraRuns.length - 1] = {
                      ...last,
                      options: { ...last.options, breakLine: true }
                    };
                  }
                  return paraRuns;
                });
              } else {
                textArray = cell.runs.map((r) => {
                  if (r.breakLine) {
                    return { text: "", options: { breakLine: true } };
                  }
                  return runToTextProp(r);
                });
              }
              return {
                text: textArray,
                options: cellOpts2
              };
            }
            const cellOpts = {
              bold: cell.isHeader || cell.style.fontWeight >= 600,
              color: rgbToHex(cell.style.color),
              fontSize: (() => {
                const base = scaledFontSizePx(
                  cell.text,
                  cell.style.fontSize,
                  cell.style.fontFamily
                );
                const columnWidth = el.colWidths?.[cellIndex] ?? el.width / row.cells.length;
                const availableWidth = Math.max(
                  1,
                  columnWidth - (cell.style.paddingLeft ?? 0) - (cell.style.paddingRight ?? 0)
                );
                const estimated = Array.from(cell.text).length * base * (japaneseTextPattern2.test(cell.text) ? 1 : 0.58);
                return pxToPoints(
                  base * Math.max(0.9, Math.min(1, availableWidth / estimated))
                );
              })(),
              fontFace: cleanFontFamily(cell.style.fontFamily, cell.text),
              align: cell.style.textAlign,
              ...cellMargin ? { margin: cellMargin } : {},
              ...cell.colspan && cell.colspan > 1 ? { colspan: cell.colspan } : {},
              ...cell.rowspan && cell.rowspan > 1 ? { rowspan: cell.rowspan } : {}
            };
            if (!isTransparent(cell.style.backgroundColor)) {
              cellOpts.fill = { color: rgbToHex(cell.style.backgroundColor) };
            }
            if (cell.style.borderColor && !isTransparent(cell.style.borderColor)) {
              cellOpts.border = {
                pt: pxToPoints(cell.style.borderWidth ?? 0) || 1,
                color: rgbToHex(cell.style.borderColor),
                ...cssBorderStyleToDash(cell.style.borderStyle) ? { dashType: cssBorderStyleToDash(cell.style.borderStyle) } : {}
              };
            }
            return { text: sanitizeText(cell.text), options: cellOpts };
          })
        ),
        {
          x,
          y,
          w,
          autoPage: false,
          // Preserve HTML column proportions when per-column widths are available
          ...el.colWidths && el.colWidths.length > 0 && el.colWidths.every((cw) => cw > 0) ? {
            // Add a proportional per-column slack to absorb PPTX/Chrome
            // font-metric variance.  DirectWrite (PPTX) renders bold text
            // slightly wider than Chrome's Skia.  A fixed absolute slack
            // was insufficient for longer header strings (e.g. "Column 2
            // (center-aligned)") where the absolute pixel variance at the
            // DirectWrite level can exceed 8 px.  Scaling each column to
            // 105% of the browser-measured width covers the observed ~3%
            // variance across all header cell lengths and font sizes while
            // adding a manageable ~5% total table width overhead.
            // Guard: if scaled column total would exceed the table width,
            // scale back proportionally so the table fits within the slide.
            colW: (() => {
              const measuredSum = el.colWidths.reduce((a, b) => a + b, 0);
              const targetWidth = Math.max(measuredSum, 1);
              return el.colWidths.map(
                (cw) => pxToInches(cw / measuredSum * targetWidth)
              );
            })()
          } : {},
          // Cell margin derived from CSS padding of the first cell.
          // Per-cell margins (set above) override this when available.
          // Fallback matches Marp's default table cell CSS padding.
          margin: tableFallbackMargin,
          ...el.rows.every((row) => (row.height ?? 0) > 0) ? {
            rowH: el.rows.map(
              (row) => pxToInches(row.height)
            )
          } : {}
        }
      );
      break;
    }
    case "code": {
      const hasCodeBg = !isTransparent(el.style.backgroundColor);
      const codeShapeOpts = {
        x,
        y,
        w,
        h,
        margin: computeTextInset(el.style),
        valign: "top",
        paraSpaceBefore: 0,
        paraSpaceAfter: 0,
        autoFit: false,
        wrap: false,
        ...hasCodeBg ? {
          shape: "rect",
          fill: { color: rgbToHex(el.style.backgroundColor) }
        } : {}
      };
      const LINE_SPACING = rendererProfile.code.lineSpacing;
      const codeNumLines = el.runs ? el.runs.filter((r) => r.breakLine).length + 1 : el.text?.split("\n").length ?? 1;
      const shapeHeightPt = h * 72;
      const baseFontSizePt = pxToPoints(el.style.fontSize);
      const maxFontSizePt = codeNumLines > 1 ? shapeHeightPt / (codeNumLines * LINE_SPACING) : shapeHeightPt;
      const codeFontScale = baseFontSizePt > maxFontSizePt ? maxFontSizePt / baseFontSizePt : 1;
      codeShapeOpts.lineSpacingMultiple = computeLineSpacing(
        el.style,
        codeNumLines > 1
      );
      const codeFontFace = cleanFontFamily(
        el.style.fontFamily,
        el.text
      );
      if (el.runs && el.runs.length > 0) {
        slide.addText(
          el.runs.map((r) => {
            if (r.breakLine) {
              return { text: "", options: { breakLine: true } };
            }
            return {
              text: sanitizeText(r.text),
              options: {
                color: rgbToHex(r.color),
                fontSize: pxToPoints(r.fontSize ?? el.style.fontSize) * codeFontScale,
                fontFace: codeFontFace,
                bold: r.bold,
                italic: r.italic
              }
            };
          }),
          codeShapeOpts
        );
      } else {
        slide.addText(sanitizeText(el.text), {
          ...codeShapeOpts,
          fontFace: codeFontFace,
          fontSize: pxToPoints(el.style.fontSize) * codeFontScale,
          color: rgbToHex(el.style.color)
        });
      }
      break;
    }
    case "image": {
      const imgOpts = {
        x,
        y,
        w,
        h,
        ...resolveImageSource(el.src)
      };
      slide.addImage(imgOpts);
      break;
    }
    case "container": {
      const bg = el.style?.backgroundColor;
      const borderWidth = el.style?.borderWidth ?? 0;
      const borderColor = el.style?.borderColor;
      const borderRadius = el.style?.borderRadius ?? 0;
      const borderLeft = el.style?.borderLeft;
      const hasBoxShadow = el.style?.boxShadow === true;
      const hasBackground = !isTransparent(bg);
      const preciseBorders = el.style?.borders;
      const borderEntries = preciseBorders ? [
        preciseBorders.top,
        preciseBorders.right,
        preciseBorders.bottom,
        preciseBorders.left
      ].filter((border) => border !== void 0) : [];
      const hasPreciseBorders = borderEntries.length > 0;
      const hasUniformRoundedBorders = hasPreciseBorders && borderRadius > 0 && borderEntries.length === 4 && borderEntries.every(
        (border) => border.width === borderEntries[0].width && border.color === borderEntries[0].color && border.style === borderEntries[0].style
      );
      const hasBorder = hasPreciseBorders || borderWidth > 0 && !!borderColor && !isTransparent(borderColor);
      const borderDashType = cssBorderStyleToDash(
        hasUniformRoundedBorders ? preciseBorders?.top?.style : el.style?.borderStyle
      );
      const lineStyle = hasBorder && (!hasPreciseBorders || hasUniformRoundedBorders) ? {
        color: rgbToHex(
          hasUniformRoundedBorders ? preciseBorders.top.color : borderColor
        ),
        width: pxToPoints(
          hasUniformRoundedBorders ? preciseBorders.top.width : borderWidth
        ),
        ...borderDashType ? { dashType: borderDashType } : {}
      } : hasBoxShadow && !hasPreciseBorders ? { color: "CCCCCC", width: 0.5 } : void 0;
      const shapeType = borderRadius > 0 ? "roundRect" : "rect";
      const minDim = Math.min(el.width, el.height);
      const rectRadius = borderRadius > 0 ? Math.min(0.5, borderRadius / (minDim / 2)) : void 0;
      const isVisibleShape = hasBackground || hasBorder || hasBoxShadow;
      const spatialText = containerAssoc?.get(
        el
      );
      const useEmbeddedPath = isVisibleShape && (spatialText && spatialText.length > 0 || isSimpleTextContainer2(el));
      if (useEmbeddedPath) {
        let runs;
        let margin;
        let firstStyle;
        if (spatialText && spatialText.length > 0) {
          runs = buildRunsFromElements(
            spatialText,
            hasBackground ? bg : void 0,
            slideBg,
            visualBgMayBeDark
          );
          margin = computeInsetFromElements(
            el,
            spatialText
          );
          const firstEl = spatialText[0];
          firstStyle = "style" in firstEl ? firstEl.style : void 0;
        } else {
          runs = buildContainerEmbeddedRuns(
            el,
            hasBackground ? bg : void 0,
            slideBg,
            visualBgMayBeDark
          );
          margin = computeContainerInset(el);
          const firstChild = el.children[0];
          firstStyle = "style" in firstChild ? firstChild.style : void 0;
        }
        const borderLeft2 = el.style?.borderLeft;
        if (borderLeft2 && borderLeft2.width > 0 && !hasPreciseBorders) {
          const extraLeft = borderLeft2.width * 0.75 + 2;
          if (Array.isArray(margin)) {
            margin = [margin[0] + extraLeft, margin[1], margin[2], margin[3]];
          } else {
            margin = [extraLeft, 0, 0, 0];
          }
        }
        const isBadgeLike = el.width <= 80 && el.height <= 80 && borderRadius >= 50;
        slide.addText(runs, {
          shape: shapeType,
          x,
          y,
          w,
          h,
          fill: hasBackground ? { color: rgbToHex(bg) } : { type: "none" },
          ...lineStyle ? { line: lineStyle } : {},
          ...rectRadius !== void 0 ? { rectRadius } : {},
          margin,
          valign: isBadgeLike ? "middle" : "top",
          align: isBadgeLike ? "center" : firstStyle?.textAlign,
          lineSpacingMultiple: isBadgeLike ? 1 : firstStyle ? computeLineSpacing(firstStyle) : void 0,
          paraSpaceBefore: 0,
          paraSpaceAfter: 0,
          charSpacing: firstStyle ? computeCharSpacing(firstStyle) : void 0,
          autoFit: false,
          wrap: true
        });
        if (borderLeft2 && borderLeft2.width > 0 && !hasPreciseBorders) {
          const bw = pxToInches(borderLeft2.width);
          const blColor = rgbToHex(borderLeft2.color);
          slide.addShape("rect", {
            x,
            y,
            w: bw,
            h,
            fill: { color: blColor },
            line: { color: blColor, width: 0.25 }
          });
        }
        if (hasPreciseBorders && !hasUniformRoundedBorders)
          drawPreciseContainerBorders(slide, preciseBorders, x, y, w, h);
        break;
      }
      const hasBadgeRuns = el.runs && el.runs.length > 0 && el.runs.some((r) => !r.breakLine && r.text.trim() !== "");
      if (hasBadgeRuns && isVisibleShape) {
        const measuredTextLines = el.textLines ?? [];
        const usesMeasuredText = measuredTextLines.length > 0 || !!el.textBounds;
        const badgeFontScale = usesMeasuredText ? 1 : computeBadgeRunFontScale(el.runs, el.width, el.height);
        if (usesMeasuredText) {
          slide.addShape(shapeType, {
            x,
            y,
            w,
            h,
            fill: hasBackground ? { color: rgbToHex(bg) } : { type: "none" },
            ...lineStyle ? { line: lineStyle } : {},
            ...rectRadius !== void 0 ? { rectRadius } : {}
          });
        }
        const renderText = (runs, bounds, lineHeight, textAlign) => {
          const textRuns = runs.map((run2) => {
            const textProps = toTP(run2);
            if (!run2.breakLine && textProps.options?.fontSize && badgeFontScale < 1) {
              textProps.options.fontSize *= badgeFontScale;
            }
            return textProps;
          });
          const measured = usesMeasuredText;
          slide.addText(
            textRuns,
            {
              ...!measured ? { shape: shapeType } : {},
              x: measured ? pxToInches(bounds.x) : x,
              y: measured ? pxToInches(bounds.y) : y,
              w: measured ? pxToInches(bounds.width) * rendererProfile.layout.nowrapWidthScale : w,
              h: measured ? pxToInches(bounds.height) : h,
              ...!measured ? {
                fill: hasBackground ? { color: rgbToHex(bg) } : { type: "none" }
              } : { fill: { type: "none" } },
              ...!measured && lineStyle ? { line: lineStyle } : {},
              ...!measured && rectRadius !== void 0 ? { rectRadius } : {},
              margin: 0,
              valign: measured ? "top" : "middle",
              align: measured ? "left" : "center",
              ...lineHeight ? { lineSpacing: pxToPoints(lineHeight) } : { lineSpacingMultiple: 1 },
              paraSpaceBefore: 0,
              paraSpaceAfter: 0
            }
          );
        };
        if (measuredTextLines.length > 0) {
          for (const line of measuredTextLines) {
            renderText(
              line.runs,
              line,
              line.lineHeight,
              line.textAlign
            );
          }
        } else {
          renderText(
            el.runs,
            el.textBounds ?? el,
            el.textBounds?.lineHeight,
            "left"
          );
        }
        if (hasPreciseBorders && !hasUniformRoundedBorders)
          drawPreciseContainerBorders(slide, preciseBorders, x, y, w, h);
        break;
      }
      if (isVisibleShape && (!hasPreciseBorders || hasBackground || hasBoxShadow)) {
        slide.addShape(shapeType, {
          x,
          y,
          w,
          h,
          fill: hasBackground ? { color: rgbToHex(bg) } : { type: "none" },
          ...lineStyle ? { line: lineStyle } : {},
          ...rectRadius !== void 0 ? { rectRadius } : {}
        });
      }
      if (borderLeft && borderLeft.width > 0 && !hasPreciseBorders) {
        const bw = pxToInches(borderLeft.width);
        slide.addShape("rect", {
          x,
          y,
          w: bw,
          h,
          fill: { color: rgbToHex(borderLeft.color) },
          line: { color: rgbToHex(borderLeft.color) }
        });
      }
      const borderBottom = el.style?.borderBottom;
      if (borderBottom && borderBottom.width > 0 && !hasPreciseBorders) {
        const bbh = pxToInches(borderBottom.width);
        const bbColor = rgbToHex(borderBottom.color);
        const bbDash = cssBorderStyleToDash(borderBottom.style);
        slide.addShape("rect", {
          x,
          y: y + h,
          w,
          h: bbh,
          // For dashed/dotted borders: omit fill so PptxGenJS generates
          // <a:noFill/>.  Passing fill:{type:'none'} is a truthy object and
          // routes through genXmlColorSelection() which only handles 'solid' —
          // it outputs nothing, leaving the shape with the slide-theme default
          // fill (potentially opaque) which would mask the dash pattern.
          // Omitting fill makes options.fill falsy → PptxGenJS emits <a:noFill/>.
          // For solid borders, a filled rect renders cleanly as a solid rule.
          ...bbDash ? {
            line: {
              color: bbColor,
              width: Math.max(0.25, pxToPoints(borderBottom.width)),
              dashType: bbDash
            }
          } : {
            fill: { color: bbColor },
            line: { color: bbColor, width: 0.25 }
          }
        });
      }
      if (hasPreciseBorders && !hasUniformRoundedBorders)
        drawPreciseContainerBorders(slide, preciseBorders, x, y, w, h);
      const childElements = el.children ?? [];
      const childAssoc = containerAssoc ?? associateContainerText2(childElements, slideW, slideH);
      const childEmbedded = new Set(
        Array.from(childAssoc.values()).flat()
      );
      for (const child of childElements) {
        if (childEmbedded.has(child)) continue;
        placeElement(
          slide,
          child,
          slideW,
          slideH,
          slideBg,
          visualBgMayBeDark,
          childAssoc
        );
      }
      break;
    }
  }
}
function placeGroupedTextElements(slide, group, slideW, slideH, slideBg, visualBgMayBeDark) {
  const minX = Math.min(...group.map((e) => e.x));
  const minY = Math.min(...group.map((e) => e.y));
  const maxX = Math.max(...group.map((e) => e.x + e.width));
  const maxY = Math.max(...group.map((e) => e.y + e.height));
  const x = pxToInches(minX);
  const firstStyle = "style" in group[0] ? group[0].style : void 0;
  const groupHalfLeading = firstStyle && firstStyle.lineHeight > 0 && firstStyle.fontSize > 0 ? (firstStyle.lineHeight - firstStyle.fontSize) / 2 : 0;
  const y = pxToInches(minY - groupHalfLeading);
  const w = pxToInches(maxX - minX);
  const h = slideH > 0 ? Math.min(pxToInches(maxY - minY), Math.max(0.01, pxToInches(slideH) - y)) : pxToInches(maxY - minY);
  const toTP = (r) => toTextProps(r, slideBg, visualBgMayBeDark);
  const allRuns = [];
  for (let i = 0; i < group.length; i++) {
    const el = group[i];
    if (i > 0) {
      allRuns.push({ text: "", options: { breakLine: true } });
    }
    if (el.type === "list") {
      const listEl = el;
      let topLevelCount = 0;
      const listRuns = listEl.items.flatMap((item, idx) => {
        const startNum = item.level === 0 && listEl.ordered ? (listEl.startNumber ?? 1) + topLevelCount : void 0;
        if (item.level === 0) topLevelCount++;
        return toListTextProps(
          item,
          listEl.ordered,
          idx < listEl.items.length - 1,
          slideBg,
          visualBgMayBeDark,
          startNum,
          listEl.listStyleType
        );
      });
      allRuns.push(...listRuns);
    } else if ("runs" in el && Array.isArray(el.runs)) {
      allRuns.push(...el.runs.map(toTP));
    }
  }
  const first = group[0];
  const style = "style" in first ? first.style : void 0;
  slide.addText(allRuns, {
    x,
    y,
    w,
    h,
    margin: style ? computeTextInset(style) : 0,
    valign: "top",
    align: style?.textAlign,
    lineSpacingMultiple: style ? computeLineSpacing(style) : void 0,
    paraSpaceBefore: 0,
    paraSpaceAfter: 0,
    charSpacing: style ? computeCharSpacing(style) : void 0
  });
}
function toTextProps(run2, slideBg = "rgb(255, 255, 255)", visualBgMayBeDark = false) {
  if (run2.breakLine) {
    return { text: "", options: { breakLine: true } };
  }
  const text = sanitizeText(run2.text);
  const highlight = computeHighlight(
    run2.backgroundColor,
    run2.color,
    slideBg,
    visualBgMayBeDark
  );
  return {
    text,
    options: {
      color: rgbToHex(run2.color),
      fontSize: pxToPoints(
        scaledFontSizePx(text, run2.fontSize ?? 16, run2.fontFamily)
      ),
      fontFace: cleanFontFamily(run2.fontFamily, run2.text),
      bold: run2.bold,
      italic: run2.italic,
      underline: run2.underline ? { style: "sng" } : void 0,
      strike: run2.strikethrough ? "sngStrike" : void 0,
      subscript: run2.subscript || void 0,
      superscript: run2.superscript || void 0,
      hyperlink: run2.hyperlink ? { url: run2.hyperlink } : void 0,
      highlight
    }
  };
}
function toListTextProps(item, ordered = false, breakAfter = false, slideBg = "rgb(255, 255, 255)", visualBgMayBeDark = false, startNumber, listStyleType) {
  const bulletOption = createListBulletOption(item, ordered, false, startNumber, listStyleType);
  if (item.runs.length === 0) {
    return [
      {
        text: sanitizeText(item.text) || " ",
        options: {
          bullet: bulletOption,
          indentLevel: item.level,
          breakLine: breakAfter
        }
      }
    ];
  }
  const groups = [[]];
  for (const run2 of item.runs) {
    if (run2.breakLine) {
      groups.push([]);
    } else {
      groups[groups.length - 1].push(run2);
    }
  }
  const result = [];
  for (let g = 0; g < groups.length; g++) {
    const group = groups[g];
    if (group.length === 0) continue;
    const isContinuation = g > 0;
    const isLastGroup = g === groups.length - 1;
    const groupBullet = isContinuation ? createListBulletOption(item, ordered, true, void 0, listStyleType) : bulletOption;
    for (let r = 0; r < group.length; r++) {
      const run2 = group[r];
      const isLastRun = r === group.length - 1;
      const needsBreakLine = isLastRun && (!isLastGroup || breakAfter);
      result.push({
        text: sanitizeText(run2.text),
        options: {
          // Always propagate bullet and indentLevel to every run in the group.
          // PptxGenJS v4.x emits <a:pPr> for each TextProp in the same
          // paragraph. LibreOffice uses the *last* <a:pPr>, so without
          // propagation the last run's pPr resets the bullet with <a:buNone/>.
          // PowerPoint uses the *first* <a:pPr> and is unaffected by this change.
          bullet: groupBullet,
          indentLevel: item.level,
          ...needsBreakLine ? { breakLine: true } : {},
          color: rgbToHex(run2.color),
          fontSize: pxToPoints(
            scaledFontSizePx(
              run2.text,
              run2.fontSize ?? 16,
              run2.fontFamily
            ) * (japaneseTextPattern2.test(run2.text) ? rendererProfile.layout.listCjkFontSizeScale / rendererProfile.text.cjkFontSizeScale : 1)
          ),
          fontFace: cleanFontFamily(run2.fontFamily, run2.text),
          bold: run2.bold,
          italic: run2.italic,
          subscript: run2.subscript || void 0,
          superscript: run2.superscript || void 0,
          highlight: computeHighlight(
            run2.backgroundColor,
            run2.color,
            slideBg,
            visualBgMayBeDark
          )
        }
      });
    }
  }
  return result;
}

// node_modules/marp-to-editable-pptx/src/abort.ts
function abortReason(signal) {
  const reason = signal.reason;
  if (reason instanceof Error || reason instanceof DOMException) return reason;
  return new DOMException("The operation was cancelled", "AbortError");
}
async function waitWithCancellation(operation, signal) {
  signal?.throwIfAborted();
  const pending = Promise.resolve(operation);
  if (!signal) return pending;
  return await new Promise((resolve, reject) => {
    const abort = () => reject(abortReason(signal));
    signal.addEventListener("abort", abort, { once: true });
    pending.then(resolve, reject).finally(() => {
      signal.removeEventListener("abort", abort);
    });
  });
}

// node_modules/marp-to-editable-pptx/src/native-pptx/index.ts
function escapeRegExp(value) {
  return value.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}
function shouldHidePaginationPseudoText(rawContent, paginationValue, paginationTotal) {
  const normalizedPaginationValue = paginationValue?.trim();
  if (!normalizedPaginationValue) return false;
  const normalizedContent = rawContent?.trim();
  if (!normalizedContent || normalizedContent === "none" || normalizedContent === "normal") {
    return false;
  }
  const strippedContent = normalizedContent.replace(/^['"]|['"]$/g, "").trim();
  if (!strippedContent) return false;
  const escapedPaginationValue = escapeRegExp(normalizedPaginationValue);
  const normalizedPaginationTotal = paginationTotal?.trim();
  const exactValuePattern = new RegExp(`^\\(?0*${escapedPaginationValue}\\)?$`);
  if (exactValuePattern.test(strippedContent)) return true;
  const labeledValuePattern = new RegExp(
    `^(?:page|slide|p\\.?|#)\\s*0*${escapedPaginationValue}$`,
    "i"
  );
  if (labeledValuePattern.test(strippedContent)) return true;
  if (!normalizedPaginationTotal) return false;
  const escapedPaginationTotal = escapeRegExp(normalizedPaginationTotal);
  const pagedFractionPattern = new RegExp(
    `^(?:(?:page|slide|p\\.?|#)\\s*)?0*${escapedPaginationValue}\\s*(?:/|of)\\s*0*${escapedPaginationTotal}$`,
    "i"
  );
  return pagedFractionPattern.test(strippedContent);
}
async function generateNativePptx(opts) {
  const { htmlPath, browserPath, width = 1280, height = 720 } = opts;
  let browser;
  let activeTiming;
  const finishTiming = () => {
    if (!activeTiming) return;
    opts.onTiming?.(
      activeTiming.stage,
      import_node_perf_hooks.performance.now() - activeTiming.startedAt
    );
    activeTiming = void 0;
  };
  const reportProgress = (message) => {
    opts.signal?.throwIfAborted();
    finishTiming();
    opts.onProgress?.(message);
    opts.signal?.throwIfAborted();
    activeTiming = { stage: message, startedAt: import_node_perf_hooks.performance.now() };
  };
  try {
    reportProgress("Launching browser");
    browser = await import_puppeteer_core.default.launch({
      executablePath: browserPath,
      headless: true,
      args: [
        "--no-sandbox",
        "--disable-setuid-sandbox",
        "--disable-dev-shm-usage",
        "--disable-gpu"
      ]
    });
    const page = await browser.newPage();
    await page.setViewport({ width, height });
    const fileUrl = (0, import_node_url2.pathToFileURL)(htmlPath).href;
    reportProgress("Loading rendered Marp HTML");
    await waitWithCancellation(
      page.goto(fileUrl, { waitUntil: "networkidle0" }),
      opts.signal
    );
    reportProgress("Waiting for slide rendering");
    await waitForRenderReady(page, opts.signal);
    reportProgress("Extracting slide structure");
    await page.addStyleTag({
      content: '*,*::before,*::after{animation:none!important;transition:none!important}[data-bespoke-marp-fragment="inactive"]{visibility:visible!important;opacity:1!important}'
    });
    await page.addStyleTag({
      content: ".bespoke-marp-osc,[data-bespoke-marp-osc],.bespoke-marp-note{display:none!important}"
    });
    await page.addScriptTag({ content: DOM_WALKER_SCRIPT });
    const slides = await page.evaluate(
      () => globalThis.extractSlides()
    );
    reportProgress(`Extracted ${slides.length} slides`);
    if (slides.some((slide) => slide.sourceHasPagination)) {
      const paginationMatches = await page.evaluate(
        () => Array.from(
          document.querySelectorAll("section[data-marpit-pagination]")
        ).map((section, index) => ({
          index,
          paginationValue: section.getAttribute("data-marpit-pagination")?.trim() ?? "",
          paginationTotal: section.getAttribute("data-marpit-pagination-total")?.trim() ?? "",
          afterContent: getComputedStyle(section, "::after").content ?? ""
        }))
      );
      const matchedPaginationIndexes = paginationMatches.filter(
        ({ afterContent, paginationValue, paginationTotal }) => shouldHidePaginationPseudoText(
          afterContent,
          paginationValue,
          paginationTotal
        )
      ).map(({ index }) => index);
      if (matchedPaginationIndexes.length > 0) {
        await page.evaluate((matchedIndexes) => {
          const sections = Array.from(
            document.querySelectorAll("section[data-marpit-pagination]")
          );
          for (const index of matchedIndexes) {
            sections[index]?.setAttribute(
              "data-native-pptx-hide-pagination-after",
              "true"
            );
          }
        }, matchedPaginationIndexes);
        await page.addStyleTag({
          content: 'section[data-native-pptx-hide-pagination-after="true"]::after{color:transparent!important;-webkit-text-fill-color:transparent!important;text-shadow:none!important}'
        });
      }
    }
    const missingUrls = await findMissingLocalUrls(slides);
    reportProgress("Resolving images and raster fallbacks");
    await rasterizeSlideTargets(
      page,
      buildBrokenContentImageJobs(slides, missingUrls),
      opts.signal
    );
    pruneMissingBackgrounds(slides, missingUrls);
    await rasterizeSlideTargets(page, buildFilteredBgJobs(slides), opts.signal);
    await rasterizeSlideTargets(
      page,
      buildCssFallbackBgJobs(slides),
      opts.signal
    );
    await rasterizeSlideTargets(
      page,
      buildFilteredContentImageJobs(slides),
      opts.signal
    );
    await rasterizeSlideTargets(
      page,
      buildRasterizeImageJobs(slides),
      opts.signal
    );
    await rasterizeSlideTargets(page, buildPartialBgJobs(slides), opts.signal);
    if (opts.debugJsonPath) {
      await (0, import_promises.writeFile)(
        opts.debugJsonPath,
        JSON.stringify(slides, null, 2),
        "utf-8"
      );
    }
    await resolveImageUrls(slides, opts.signal);
    reportProgress("Building editable PowerPoint");
    const renderPlan = createPptxRenderPlan(slides);
    if (opts.debugJsonPath) {
      await (0, import_promises.writeFile)(
        opts.debugJsonPath.replace(/\.json$/i, ".plan.json"),
        JSON.stringify(
          {
            totals: renderPlan.totals,
            slides: renderPlan.slides.map((slide) => ({
              background: slide.background,
              visualBgMayBeDark: slide.visualBgMayBeDark,
              editableElementCount: slide.editableElementCount,
              rasterizedElementCount: slide.rasterizedElementCount,
              editableAreaRatio: slide.editableAreaRatio
            }))
          },
          null,
          2
        ),
        "utf-8"
      );
    }
    const pptx = executePptxRenderPlan(renderPlan);
    reportProgress("Writing PowerPoint output");
    const output = await waitWithCancellation(
      pptx.write({ outputType: "nodebuffer" }),
      opts.signal
    );
    opts.signal?.throwIfAborted();
    finishTiming();
    return Buffer.from(output);
  } finally {
    finishTiming();
    await browser?.close();
  }
}
var RENDER_READY_TIMEOUT_MS = 1e4;
var RENDER_STABLE_MS = 300;
async function waitForRenderReady(page, signal) {
  await waitWithCancellation(page.waitForFunction(
    (stableMs) => {
      if (document.readyState !== "complete") return false;
      if (document.fonts && document.fonts.status !== "loaded") return false;
      const pendingImages = Array.from(document.images).some(
        (image) => !image.complete
      );
      if (pendingImages) return false;
      const sections = Array.from(
        document.querySelectorAll("svg[data-marpit-svg]")
      );
      const signature = sections.map((svg) => {
        const rect = svg.getBoundingClientRect();
        return [
          Math.round(rect.left),
          Math.round(rect.top),
          Math.round(rect.width),
          Math.round(rect.height),
          svg.querySelectorAll("*").length,
          svg.querySelectorAll("svg").length,
          svg.textContent?.length ?? 0
        ].join(":");
      }).join("|");
      const runtime = globalThis;
      const now = globalThis.performance.now();
      if (runtime.__nativePptxRenderState?.signature !== signature) {
        runtime.__nativePptxRenderState = { signature, stableSince: now };
        return false;
      }
      return now - runtime.__nativePptxRenderState.stableSince >= stableMs;
    },
    { polling: 50, timeout: RENDER_READY_TIMEOUT_MS },
    RENDER_STABLE_MS
  ), signal);
}
async function rasterizeSlideTargets(page, jobs, signal) {
  if (jobs.length === 0) return;
  for (const { slideIdx, targets, setup, teardown } of jobs) {
    signal?.throwIfAborted();
    await page.evaluate((n) => {
      window.location.hash = "#" + n;
    }, slideIdx + 1);
    await waitForRenderReady(page, signal);
    let slideOriginX = 0;
    let slideOriginY = 0;
    if (targets.some((t) => t.slideRelative)) {
      const origin = await page.evaluate((n) => {
        const target = document.getElementById(String(n));
        if (target) {
          const fo = target.parentElement;
          const svg = fo?.tagName.toLowerCase() === "foreignobject" ? fo.parentElement : null;
          const ref = svg?.hasAttribute("data-marpit-svg") ? svg : target;
          const r = ref.getBoundingClientRect();
          return {
            x: r.left + window.scrollX,
            y: r.top + window.scrollY
          };
        }
        const visibleSections = Array.from(
          document.querySelectorAll("section")
        ).map((section) => {
          const rect = section.getBoundingClientRect();
          const visibleWidth = Math.min(rect.right, window.innerWidth) - Math.max(rect.left, 0);
          const visibleHeight = Math.min(rect.bottom, window.innerHeight) - Math.max(rect.top, 0);
          const visibleArea = Math.max(0, visibleWidth) * Math.max(0, visibleHeight);
          return { rect, visibleArea };
        }).filter((entry) => entry.visibleArea > 0);
        if (visibleSections.length === 0) {
          console.warn(
            `[rasterize] slide ${n}: section id not found and no visible sections in viewport; using origin (0,0) which may produce an incorrect clip for bespoke.js HTML layouts.`
          );
          return { x: 0, y: 0 };
        }
        visibleSections.sort((left, right) => right.visibleArea - left.visibleArea);
        return {
          x: visibleSections[0].rect.left + window.scrollX,
          y: visibleSections[0].rect.top + window.scrollY
        };
      }, slideIdx + 1);
      slideOriginX = origin.x;
      slideOriginY = origin.y;
    }
    let captureError;
    try {
      if (setup) await setup(page, slideIdx);
      for (const { clip, slideRelative, onCapture } of targets) {
        signal?.throwIfAborted();
        const effectiveClip = slideRelative ? {
          x: Math.round(slideOriginX + clip.x),
          y: Math.round(slideOriginY + clip.y),
          width: clip.width,
          height: clip.height
        } : clip;
        if (effectiveClip.width <= 0 || effectiveClip.height <= 0) continue;
        try {
          const raw = await waitWithCancellation(
            page.screenshot({
              type: "png",
              clip: effectiveClip
            }),
            signal
          );
          onCapture(
            "data:image/png;base64," + Buffer.from(raw).toString("base64")
          );
        } catch (error) {
          if (signal?.aborted) throw error;
          captureError ??= new Error(
            `Failed to rasterize slide ${slideIdx + 1}: ` + String(error instanceof Error ? error.message : error)
          );
        }
      }
    } finally {
      if (teardown) await teardown(page, slideIdx);
    }
    if (captureError) throw captureError;
  }
  await page.evaluate(() => {
    window.location.hash = "#1";
  });
  await waitForRenderReady(page, signal);
}
var ADVANCED_LAYERS_SELECTOR = 'section[data-marpit-advanced-background="content"], section[data-marpit-advanced-background="pseudo"]';
async function hideAdvancedLayers(page) {
  await page.evaluate((sel) => {
    document.querySelectorAll(sel).forEach(
      (el) => el.style.setProperty(
        "visibility",
        "hidden",
        "important"
      )
    );
  }, ADVANCED_LAYERS_SELECTOR);
}
async function restoreAdvancedLayers(page) {
  await page.evaluate((sel) => {
    document.querySelectorAll(sel).forEach((el) => el.style.removeProperty("visibility"));
  }, ADVANCED_LAYERS_SELECTOR);
}
async function hideSectionChildren(page, slideIdx) {
  const found = await page.evaluate((id) => {
    const section = document.getElementById(id);
    if (!section) return false;
    Array.from(section.children).forEach(
      (el) => el.style.setProperty("visibility", "hidden", "important")
    );
    return true;
  }, String(slideIdx + 1));
  if (!found) {
    console.warn(
      `[rasterize] hideSectionChildren: section id="${slideIdx + 1}" not found; children not hidden \u2014 background screenshot may include slide content.`
    );
  }
}
async function restoreSectionChildren(page, slideIdx) {
  await page.evaluate((id) => {
    const section = document.getElementById(id);
    if (!section) return;
    Array.from(section.children).forEach(
      (el) => el.style.removeProperty("visibility")
    );
  }, String(slideIdx + 1));
}
function isLocalImagePath(url) {
  if (!url) return false;
  if (url.startsWith("data:")) return false;
  if (url.startsWith("http://") || url.startsWith("https://")) return false;
  return true;
}
async function fileUrlToDataUrl(url) {
  const filePath = url.startsWith("file:") ? (0, import_node_url2.fileURLToPath)(url) : url;
  const buf = await (0, import_promises.readFile)(filePath);
  const ext = filePath.split(".").pop()?.toLowerCase() ?? "";
  const mime = ext === "jpg" || ext === "jpeg" ? "image/jpeg" : ext === "gif" ? "image/gif" : ext === "webp" ? "image/webp" : ext === "svg" ? "image/svg+xml" : "image/png";
  return `data:${mime};base64,${buf.toString("base64")}`;
}
async function resolveImageUrls(slides, signal, converter = fileUrlToDataUrl) {
  const jobs = [];
  const converted = /* @__PURE__ */ new Map();
  const convert = (url) => {
    let result = converted.get(url);
    if (!result) {
      result = converter(url);
      converted.set(url, result);
    }
    return result;
  };
  function resolveEl(el) {
    if (el.type === "image" && isLocalImagePath(el.src)) {
      const sourceUrl = el.src;
      jobs.push(
        convert(sourceUrl).then((d) => {
          el.src = d;
        })
      );
    }
    if ("children" in el && Array.isArray(el.children)) {
      for (const child of el.children) resolveEl(child);
    }
  }
  for (const slide of slides) {
    for (const bg of slide.backgroundImages ?? []) {
      if (isLocalImagePath(bg.url)) {
        const sourceUrl = bg.url;
        jobs.push(
          convert(sourceUrl).then((d) => {
            bg.url = d;
          })
        );
      }
    }
    for (const el of slide.elements ?? []) resolveEl(el);
  }
  signal?.throwIfAborted();
  await waitWithCancellation(Promise.all(jobs).then(() => void 0), signal);
}
async function findMissingLocalUrls(slides) {
  const urlsToCheck = /* @__PURE__ */ new Set();
  function collectImageUrls(elements) {
    for (const el of elements) {
      if (el.type === "image" && isLocalImagePath(el.src))
        urlsToCheck.add(el.src);
      if ("children" in el && Array.isArray(el.children)) {
        collectImageUrls(el.children);
      }
    }
  }
  for (const slide of slides) {
    for (const bg of slide.backgroundImages ?? []) {
      if (isLocalImagePath(bg.url)) urlsToCheck.add(bg.url);
    }
    collectImageUrls(slide.elements ?? []);
  }
  const missing = /* @__PURE__ */ new Set();
  await Promise.all(
    [...urlsToCheck].map(async (url) => {
      try {
        const filePath = url.startsWith("file:") ? (0, import_node_url2.fileURLToPath)(url) : url;
        await (0, import_promises.access)(filePath);
      } catch {
        missing.add(url);
      }
    })
  );
  return missing;
}
function buildBrokenContentImageJobs(slides, missingUrls) {
  function collectBrokenImages(elements) {
    const result = [];
    for (const el of elements) {
      if (el.type === "image" && missingUrls.has(el.src)) result.push(el);
      if ("children" in el && Array.isArray(el.children)) {
        result.push(...collectBrokenImages(el.children));
      }
    }
    return result;
  }
  return slides.flatMap((s, i) => {
    const imgs = collectBrokenImages(s.elements ?? []);
    if (imgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: imgs.map(
          (img) => ({
            clip: {
              x: Math.round(img.x),
              y: Math.round(img.y),
              width: Math.round(img.width),
              height: Math.round(img.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              img.src = dataUrl;
            }
          })
        )
      }
    ];
  });
}
function pruneMissingBackgrounds(slides, missingUrls) {
  for (const slide of slides) {
    slide.backgroundImages = (slide.backgroundImages ?? []).filter(
      (bg) => !missingUrls.has(bg.url)
    );
  }
}
function buildFilteredBgJobs(slides) {
  return slides.flatMap((s, i) => {
    const bgs = (s.backgroundImages ?? []).filter((b) => b.cssFilter);
    if (bgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg) => ({
            clip: {
              x: Math.round(bg.x),
              y: Math.round(bg.y),
              width: Math.round(bg.width),
              height: Math.round(bg.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl;
              delete bg.cssFilter;
            }
          })
        ),
        setup: (p) => hideAdvancedLayers(p),
        teardown: (p) => restoreAdvancedLayers(p)
      }
    ];
  });
}
function buildCssFallbackBgJobs(slides) {
  return slides.flatMap((s, i) => {
    const bgs = (s.backgroundImages ?? []).filter((b) => b.fromCssFallback);
    if (bgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg) => ({
            clip: {
              x: 0,
              y: 0,
              width: Math.round(s.width),
              height: Math.round(s.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl;
            }
          })
        ),
        setup: (p, idx) => hideSectionChildren(p, idx),
        teardown: (p, idx) => restoreSectionChildren(p, idx)
      }
    ];
  });
}
function collectFilteredContentImages(elements) {
  const result = [];
  for (const el of elements ?? []) {
    if (el.type === "image" && el.cssFilter) result.push(el);
    if ("children" in el && Array.isArray(el.children)) {
      result.push(...collectFilteredContentImages(el.children));
    }
  }
  return result;
}
function buildFilteredContentImageJobs(slides) {
  return slides.flatMap((s, i) => {
    const imgs = collectFilteredContentImages(s.elements);
    if (imgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: imgs.map(
          (img) => ({
            clip: {
              x: Math.round(img.x),
              y: Math.round(img.y),
              width: Math.round(img.width),
              height: Math.round(img.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              img.src = dataUrl;
              delete img.cssFilter;
            }
          })
        )
      }
    ];
  });
}
function collectRasterizeImages(elements) {
  const result = [];
  for (const el of elements ?? []) {
    if (el.type === "image" && el.rasterize) result.push(el);
    if ("children" in el && Array.isArray(el.children)) {
      result.push(...collectRasterizeImages(el.children));
    }
  }
  return result;
}
function buildRasterizeImageJobs(slides) {
  return slides.flatMap((s, i) => {
    const imgs = collectRasterizeImages(s.elements);
    if (imgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: imgs.map(
          (img) => ({
            clip: {
              x: Math.round(img.x),
              y: Math.round(img.y),
              width: Math.round(img.width),
              height: Math.round(img.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              img.src = dataUrl;
              delete img.rasterize;
            }
          })
        )
      }
    ];
  });
}
function buildPartialBgJobs(slides) {
  return slides.flatMap((s, i) => {
    const bgs = (s.backgroundImages ?? []).filter((b) => {
      if (b.cssFilter || b.fromCssFallback) return false;
      if (b.url.startsWith("data:")) return false;
      if (b.backgroundSizeContain) return true;
      const isFullSlide = b.x <= 1 && b.y <= 1 && Math.abs(b.width - s.width) <= 2 && Math.abs(b.height - s.height) <= 2;
      return !isFullSlide;
    });
    if (bgs.length === 0) return [];
    return [
      {
        slideIdx: i,
        targets: bgs.map(
          (bg) => ({
            clip: {
              x: Math.round(bg.x),
              y: Math.round(bg.y),
              width: Math.round(bg.width),
              height: Math.round(bg.height)
            },
            slideRelative: true,
            onCapture(dataUrl) {
              bg.url = dataUrl;
            }
          })
        ),
        setup: (p) => hideAdvancedLayers(p),
        teardown: (p) => restoreAdvancedLayers(p)
      }
    ];
  });
}

// src/plan.ts
var import_node_fs2 = require("node:fs");
var import_node_path2 = __toESM(require("node:path"));
var import_cosmiconfig = require("cosmiconfig");

// src/options.ts
var import_node_util = require("node:util");
var UsageError = class extends Error {
};
var DEFAULT_PARALLEL = 5;
var USAGE = `Usage:
  marp-to-editable-pptx-cli [options] <files...> [-- marp-cli-args...]
  marp-to-editable-pptx-cli [options] -I <dir>

Convert Marp Markdown into editable PowerPoint (.pptx) files. Options follow
Marp CLI ("marp --pptx --pptx-editable ..."), so the same command lines work;
PPTX output is always editable, and --pptx / --pptx-editable are the defaults.

A file of "-", or no file with piped input, reads Markdown from standard input.
Each slides.md is written beside it as slides.pptx unless -o is given. Unlike
Marp CLI, existing files are never overwritten without --force or --update.

Basic options (as in Marp CLI):
  -o, --output <path>      Output file ("-" for standard output), or the
                           output directory when --input-dir is given
  -I, --input-dir <dir>    Convert every Markdown file below <dir>, keeping
                           the folder structure in the output directory
  -c, --config-file, --config <file>
                           Marp config file (default: look for .marprc.yml,
                           marp.config.js, ... in each deck's folder, then its
                           parents up to the project root)
      --no-config-file, --no-config
                           Do not look for a Marp config file
  -P, --parallel [num]     Convert up to <num> files at once (default: ${DEFAULT_PARALLEL})
      --no-parallel        Convert one file at a time
  -d, --debug              Show Marp CLI output and conversion progress
  -v, --version            Show versions
  -h, --help               Show help

Converter, browser and Marp options (as in Marp CLI):
      --pptx               Convert to PowerPoint (always on)
      --pptx-editable      Generate editable PPTX (always on)
      --allow-local-files  Allow access to local files from Markdown
      --html, --no-html    Enable or disable HTML tags in Markdown
      --theme <name|file>  Override the theme by name or CSS file
      --theme-set <path>   Additional theme CSS file or directory (repeatable)
      --browser <kind>     auto, chrome or edge (default: auto)
      --browser-path <path>
                           Browser executable (default: $CHROME_PATH, or found
                           automatically)
  -- <args...>             Pass any remaining arguments to Marp CLI's
                           Markdown-to-HTML step

Batch and overwrite options (extensions):
  -f, --force              Overwrite existing output files
  -u, --update             Only convert when the output is missing or older
                           than its input (or its config/theme files); stale
                           outputs are overwritten
  -n, --dry-run            Show what would be converted, without converting
  -k, --keep-going         Continue after an error (default: stop at the first)
  -r, --recursive          Include subfolders of directories given as <files>
      --output-dir <dir>   Write outputs for <files> into <dir>
  -q, --quiet              Only print errors
      --verbose            Same as --debug

Examples:
  marp-to-editable-pptx-cli --pptx --pptx-editable slides.md
  marp-to-editable-pptx-cli slides.md -o deck.pptx
  marp-to-editable-pptx-cli -u -I decks/ -o build/
  cat slides.md | marp-to-editable-pptx-cli -o - > deck.pptx
`;
var UNSUPPORTED_OUTPUTS = ["pdf", "notes", "image", "images"];
var UNSUPPORTED_MODES = ["watch", "server", "preview"];
function normalizeParallel(argv) {
  const out = [];
  for (let i = 0; i < argv.length; i++) {
    const arg = argv[i];
    if ((arg === "-P" || arg === "--parallel") && !/^\d+$/.test(argv[i + 1] ?? "")) {
      out.push(`--parallel=${DEFAULT_PARALLEL}`);
    } else {
      out.push(arg);
    }
  }
  return out;
}
function parseCliArgs(argv) {
  const sep = argv.indexOf("--");
  const ownArgs = normalizeParallel(sep >= 0 ? argv.slice(0, sep) : argv);
  const marpArgs = sep >= 0 ? argv.slice(sep + 1) : [];
  let parsed;
  try {
    parsed = (0, import_node_util.parseArgs)({
      args: ownArgs,
      allowPositionals: true,
      allowNegative: true,
      strict: true,
      options: {
        // Marp CLI basic options
        output: { type: "string", short: "o" },
        "input-dir": { type: "string", short: "I" },
        "config-file": { type: "string", short: "c" },
        config: { type: "string" },
        // Declared so the negated forms are accepted: allowNegative alone
        // only applies to boolean options. They are reported as `false` on
        // the positive name (e.g. `config: false`).
        "no-config-file": { type: "boolean" },
        "no-config": { type: "boolean" },
        parallel: { type: "string", short: "P" },
        "no-parallel": { type: "boolean" },
        debug: { type: "boolean", short: "d" },
        version: { type: "boolean", short: "v" },
        help: { type: "boolean", short: "h" },
        // Marp CLI converter / Marp / browser options
        pptx: { type: "boolean" },
        "pptx-editable": { type: "boolean" },
        "allow-local-files": { type: "boolean" },
        html: { type: "boolean" },
        theme: { type: "string" },
        "theme-set": { type: "string", multiple: true },
        browser: { type: "string" },
        "browser-path": { type: "string" },
        // Marp CLI options that are recognised only to be rejected clearly
        pdf: { type: "boolean" },
        notes: { type: "boolean" },
        image: { type: "string" },
        images: { type: "string" },
        watch: { type: "boolean", short: "w" },
        server: { type: "boolean", short: "s" },
        preview: { type: "boolean", short: "p" },
        // Extensions
        force: { type: "boolean", short: "f" },
        update: { type: "boolean", short: "u" },
        "dry-run": { type: "boolean", short: "n" },
        "keep-going": { type: "boolean", short: "k" },
        recursive: { type: "boolean", short: "r" },
        "output-dir": { type: "string" },
        quiet: { type: "boolean", short: "q" },
        verbose: { type: "boolean" }
      }
    });
  } catch (error) {
    throw new UsageError(error instanceof Error ? error.message : String(error));
  }
  const { values: v, positionals } = parsed;
  for (const name of UNSUPPORTED_OUTPUTS) {
    if (v[name] !== void 0 && v[name] !== false) {
      throw new UsageError(
        `--${name} is not supported: this tool only produces editable PPTX (use Marp CLI for other formats)`
      );
    }
  }
  if (v.pptx === false || v["pptx-editable"] === false) {
    throw new UsageError(
      `--no-${v.pptx === false ? "pptx" : "pptx-editable"} is not supported: this tool only produces editable PPTX (use Marp CLI for image-based PPTX)`
    );
  }
  for (const name of UNSUPPORTED_MODES) {
    if (v[name]) throw new UsageError(`--${name} is not supported by this tool`);
  }
  const configValues = [v["config-file"], v.config];
  const configs = configValues.filter((c) => typeof c === "string");
  const noConfig = configValues.includes(false);
  if (configs.length > 1) throw new UsageError("give the config file only once");
  if (configs.length > 0 && noConfig) {
    throw new UsageError("--config-file and --no-config-file cannot be used together");
  }
  const parallelValue = v.parallel;
  let jobs = DEFAULT_PARALLEL;
  if (parallelValue === false) jobs = 1;
  else if (typeof parallelValue === "string") {
    jobs = Number(parallelValue);
    if (!Number.isInteger(jobs) || jobs < 0) {
      throw new UsageError(`--parallel must be a whole number (got "${parallelValue}")`);
    }
    jobs = Math.max(1, jobs);
  }
  const browser = v.browser ?? "auto";
  if (!["auto", "chrome", "edge"].includes(browser)) {
    throw new UsageError(
      `--browser must be one of auto, chrome, edge (got "${browser}")`
    );
  }
  const inputDir = v["input-dir"];
  let inputs = positionals;
  let output = v.output;
  let outputDir = v["output-dir"];
  if (inputDir !== void 0) {
    if (positionals.length > 0) {
      throw new UsageError("give either --input-dir or input files, not both");
    }
    if (output === "-") {
      throw new UsageError("--input-dir cannot write to standard output");
    }
    if (output !== void 0 && outputDir !== void 0) {
      throw new UsageError("--output and --output-dir cannot be used together");
    }
    inputs = [inputDir];
    outputDir = output ?? outputDir;
    output = void 0;
  } else if (output !== void 0 && outputDir !== void 0) {
    throw new UsageError("--output and --output-dir cannot be used together");
  }
  const opts = {
    inputs,
    output,
    outputDir,
    force: !!v.force,
    update: !!v.update,
    dryRun: !!v["dry-run"],
    recursive: !!v.recursive || inputDir !== void 0,
    keepGoing: !!v["keep-going"],
    jobs,
    config: configs[0],
    noConfig,
    theme: v.theme,
    themeSet: v["theme-set"] ?? [],
    html: v.html,
    allowLocalFiles: !!v["allow-local-files"],
    browser,
    browserPath: v["browser-path"] ?? (process.env.CHROME_PATH || void 0),
    marpArgs,
    quiet: !!v.quiet,
    verbose: !!v.debug || !!v.verbose,
    help: !!v.help,
    version: !!v.version
  };
  if (opts.quiet && opts.verbose) {
    throw new UsageError("--quiet and --debug/--verbose cannot be used together");
  }
  return opts;
}

// src/plan.ts
var STDIO = "-";
var MARKDOWN_EXTENSIONS = /* @__PURE__ */ new Set([".md", ".markdown"]);
var SKIPPED_DIRECTORIES = /* @__PURE__ */ new Set(["node_modules"]);
function tryStat(p) {
  try {
    return (0, import_node_fs2.statSync)(p);
  } catch {
    return void 0;
  }
}
function isMarkdownFile(name) {
  return MARKDOWN_EXTENSIONS.has(import_node_path2.default.extname(name).toLowerCase());
}
function findMarkdownFiles(dir, recursive) {
  const found = [];
  const entries = (0, import_node_fs2.readdirSync)(dir, { withFileTypes: true }).sort(
    (a, b) => a.name.localeCompare(b.name)
  );
  for (const entry of entries) {
    if (entry.name.startsWith(".")) continue;
    const full = import_node_path2.default.join(dir, entry.name);
    if (entry.isDirectory()) {
      if (recursive && !SKIPPED_DIRECTORIES.has(entry.name)) {
        found.push(...findMarkdownFiles(full, recursive));
      }
    } else if (isMarkdownFile(entry.name) && tryStat(full)?.isFile()) {
      found.push(full);
    }
  }
  return found;
}
function replaceExtension(file, ext) {
  const parsed = import_node_path2.default.parse(file);
  return import_node_path2.default.join(parsed.dir, parsed.name + ext);
}
function makeConfigFinder() {
  const explorer = (0, import_cosmiconfig.cosmiconfigSync)("marp", { searchStrategy: "project" });
  const cache = /* @__PURE__ */ new Map();
  return (dir) => {
    const key = import_node_path2.default.resolve(dir);
    if (!cache.has(key)) {
      const found = explorer.search(key);
      cache.set(key, found && !found.isEmpty ? found.filepath : void 0);
    }
    return cache.get(key);
  };
}
function sharedDependencyMtime(opts) {
  const paths = [opts.theme, ...opts.themeSet].filter(
    (p) => !!p
  );
  let latest = 0;
  for (const p of paths) {
    const st = tryStat(p);
    if (!st) continue;
    latest = Math.max(latest, st.mtimeMs);
    if (st.isDirectory()) {
      for (const name of (0, import_node_fs2.readdirSync)(p)) {
        if (import_node_path2.default.extname(name).toLowerCase() !== ".css") continue;
        latest = Math.max(latest, tryStat(import_node_path2.default.join(p, name))?.mtimeMs ?? 0);
      }
    }
  }
  return latest;
}
function planJobs(opts) {
  const inputs = opts.inputs.length > 0 ? opts.inputs : [STDIO];
  const sources = [];
  const seen = /* @__PURE__ */ new Set();
  let stdinCount = 0;
  for (const input of inputs) {
    if (input === STDIO) {
      if (++stdinCount > 1) {
        throw new UsageError('standard input ("-") can only be given once');
      }
      sources.push({ file: STDIO, base: "." });
      continue;
    }
    const st = tryStat(input);
    if (!st) throw new UsageError(`input not found: ${input}`);
    const add = (file, base) => {
      const key = import_node_path2.default.resolve(file);
      if (seen.has(key)) return;
      seen.add(key);
      sources.push({ file, base });
    };
    if (st.isDirectory()) {
      const files = findMarkdownFiles(input, opts.recursive);
      if (files.length === 0 && !opts.quiet) {
        process.stderr.write(`warning: no Markdown files found in ${input}
`);
      }
      for (const file of files) add(file, input);
    } else {
      add(input, import_node_path2.default.dirname(input));
    }
  }
  if (opts.output !== void 0 && sources.length !== 1) {
    throw new UsageError(
      `--output can only be used with a single input (got ${sources.length}); use -I <dir> -o <dir>, or --output-dir, for multiple inputs`
    );
  }
  const jobs = sources.map(({ file, base }) => {
    let output;
    if (opts.output !== void 0) {
      output = opts.output;
    } else if (file === STDIO) {
      throw new UsageError(
        'reading from standard input requires --output (use "-o -" to write to standard output)'
      );
    } else if (opts.outputDir !== void 0) {
      output = import_node_path2.default.join(
        opts.outputDir,
        replaceExtension(import_node_path2.default.relative(base, file), ".pptx")
      );
    } else {
      output = replaceExtension(file, ".pptx");
    }
    return { input: file, output, action: "create" };
  });
  const outputs = /* @__PURE__ */ new Map();
  for (const job of jobs) {
    if (job.output === STDIO) continue;
    const key = import_node_path2.default.resolve(job.output);
    if (job.input !== STDIO && key === import_node_path2.default.resolve(job.input)) {
      throw new UsageError(`output would overwrite its input: ${job.input}`);
    }
    const other = outputs.get(key);
    if (other !== void 0) {
      throw new UsageError(
        `${other} and ${job.input} would both be written to ${job.output}`
      );
    }
    outputs.set(key, job.input);
  }
  if (opts.config !== void 0) {
    if (!tryStat(opts.config)?.isFile()) {
      throw new UsageError(`config file not found: ${opts.config}`);
    }
    for (const job of jobs) job.config = opts.config;
  } else if (!opts.noConfig) {
    const findConfig = makeConfigFinder();
    for (const job of jobs) {
      const dir = job.input === STDIO ? process.cwd() : import_node_path2.default.dirname(job.input);
      job.config = findConfig(dir);
    }
  }
  const depMtime = opts.update ? sharedDependencyMtime(opts) : 0;
  for (const job of jobs) {
    if (job.output === STDIO) continue;
    const outStat = tryStat(job.output);
    if (!outStat) continue;
    if (outStat.isDirectory()) {
      throw new UsageError(`output path is a directory: ${job.output}`);
    }
    if (opts.update && job.input !== STDIO) {
      const inMtime = (0, import_node_fs2.statSync)(job.input).mtimeMs;
      const configMtime = job.config ? tryStat(job.config)?.mtimeMs ?? 0 : 0;
      job.action = outStat.mtimeMs >= Math.max(inMtime, depMtime, configMtime) ? "up-to-date" : "overwrite";
    } else if (opts.force || opts.update) {
      job.action = "overwrite";
    } else {
      job.action = "exists";
    }
  }
  return jobs;
}

// src/convert.ts
var pendingTempFiles = /* @__PURE__ */ new Set();
function tempName(dir, suffix) {
  const id = (0, import_node_crypto.randomBytes)(6).toString("hex");
  const file = import_node_path3.default.join(dir, `.marp-editable-pptx-${id}${suffix}`);
  pendingTempFiles.add(file);
  return file;
}
async function removeTemp(file) {
  pendingTempFiles.delete(file);
  try {
    await (0, import_promises2.unlink)(file);
  } catch {
  }
}
async function readStdin() {
  const chunks = [];
  for await (const chunk of process.stdin) chunks.push(chunk);
  return Buffer.concat(chunks);
}
function marpCliArgs(input, htmlPath, config, opts) {
  const args = [input, "-o", htmlPath];
  if (config !== void 0) args.push("--config-file", config);
  else args.push("--no-config-file");
  if (opts.theme !== void 0) args.push("--theme", opts.theme);
  for (const theme of opts.themeSet) args.push("--theme-set", theme);
  if (opts.html === true) args.push("--html");
  if (opts.html === false) args.push("--no-html");
  if (opts.allowLocalFiles) args.push("--allow-local-files");
  args.push(...opts.marpArgs);
  return args;
}
var marpQueue = Promise.resolve();
function serialized(fn) {
  const result = marpQueue.then(fn, fn);
  marpQueue = result.catch(() => {
  });
  return result;
}
async function runMarpCli(args) {
  const captured = [];
  const methods = ["log", "info", "warn", "error"];
  const originals = methods.map((m) => console[m]);
  for (const m of methods) {
    console[m] = (...data) => {
      captured.push((0, import_node_util2.format)(...data));
    };
  }
  try {
    const exitCode = await (0, import_marp_cli.marpCli)(args, {});
    return { exitCode, output: captured.join("\n") };
  } finally {
    methods.forEach((m, i) => console[m] = originals[i]);
  }
}
async function convertFile(req) {
  const { opts, signal, log } = req;
  const cleanup = [];
  try {
    let inputPath = req.input;
    if (inputPath === STDIO) {
      inputPath = tempName(process.cwd(), ".md");
      cleanup.push(inputPath);
      await (0, import_promises2.writeFile)(inputPath, await readStdin());
    }
    const htmlPath = tempName(import_node_path3.default.dirname(import_node_path3.default.resolve(inputPath)), ".html");
    cleanup.push(htmlPath);
    log("Rendering Markdown with Marp CLI");
    const { exitCode, output } = await serialized(
      () => runMarpCli(marpCliArgs(inputPath, htmlPath, req.config, opts))
    );
    if (opts.verbose && output) process.stderr.write(output + "\n");
    if (exitCode !== 0) {
      throw new Error(
        `Marp CLI exited with code ${exitCode}` + (output && !opts.verbose ? `:
${output}` : "")
      );
    }
    signal.throwIfAborted();
    const pptx = await generateNativePptx({
      htmlPath,
      browserPath: req.browserPath,
      signal,
      onProgress: log
    });
    signal.throwIfAborted();
    if (req.output === STDIO) {
      await new Promise(
        (resolve, reject) => process.stdout.write(pptx, (err) => err ? reject(err) : resolve())
      );
      return;
    }
    const outDir = import_node_path3.default.dirname(import_node_path3.default.resolve(req.output));
    await (0, import_promises2.mkdir)(outDir, { recursive: true });
    const tmpOut = tempName(outDir, ".pptx.tmp");
    cleanup.push(tmpOut);
    await (0, import_promises2.writeFile)(tmpOut, pptx);
    signal.throwIfAborted();
    await (0, import_promises2.rename)(tmpOut, req.output);
  } finally {
    await Promise.all(cleanup.map(removeTemp));
  }
}

// src/cli.ts
function removePendingTempFiles() {
  for (const file of pendingTempFiles) {
    try {
      (0, import_node_fs3.unlinkSync)(file);
    } catch {
    }
  }
  pendingTempFiles.clear();
}
var label = (p, stdio) => p === STDIO ? stdio : p;
function describe(job) {
  return `${label(job.input, "<stdin>")} -> ${label(job.output, "<stdout>")}`;
}
async function run(opts) {
  const info = (msg) => {
    if (!opts.quiet) process.stderr.write(msg + "\n");
  };
  const error = (msg) => process.stderr.write(`error: ${msg}
`);
  const jobs = planJobs(opts);
  if (jobs.length === 0) return 0;
  if (!opts.dryRun && process.stdout.isTTY && jobs.some((j) => j.output === STDIO)) {
    throw new UsageError(
      "refusing to write binary .pptx data to a terminal; redirect standard output or use -o FILE"
    );
  }
  if (opts.dryRun) {
    const verb = {
      create: "convert",
      overwrite: "convert (overwrite)",
      "up-to-date": "skip (up to date)",
      exists: "skip (output exists; use --force or --update)"
    };
    for (const job of jobs) info(`${verb[job.action]}: ${describe(job)}`);
    return 0;
  }
  const todo = jobs.filter((j) => j.action === "create" || j.action === "overwrite");
  let browserPath;
  if (todo.length > 0) {
    browserPath = detectBrowserPath(opts.browser, opts.browserPath);
    if (!browserPath) {
      error(
        opts.browserPath ? `browser executable not found: ${opts.browserPath}` : "could not find a Chromium-based browser (Chrome or Edge), which is required for PPTX export. Install one, or pass --browser-path (or set CHROME_PATH)."
      );
      return 1;
    }
  }
  process.on("exit", removePendingTempFiles);
  const controller = new AbortController();
  const onSignal = (sig) => {
    if (controller.signal.aborted) process.exit(130);
    info(`
Received ${sig}, stopping...`);
    controller.abort(new DOMException("Conversion cancelled", "AbortError"));
  };
  process.on("SIGINT", onSignal);
  process.on("SIGTERM", onSignal);
  let converted = 0;
  let skipped = 0;
  let failed = 0;
  let stopped = false;
  const multi = jobs.length > 1;
  const fail = (msg) => {
    failed++;
    error(msg);
    if (!opts.keepGoing) stopped = true;
  };
  const processJob = async (job) => {
    if (job.action === "up-to-date") {
      skipped++;
      if (opts.verbose) info(`Up to date: ${describe(job)}`);
      return;
    }
    if (job.action === "exists") {
      fail(`output exists, not overwriting (use --force or --update): ${job.output}`);
      return;
    }
    const started = Date.now();
    info(`Converting: ${describe(job)}`);
    if (opts.verbose && job.config) info(`  using Marp config ${job.config}`);
    try {
      await convertFile({
        input: job.input,
        output: job.output,
        config: job.config,
        browserPath,
        opts,
        signal: controller.signal,
        log: (msg) => {
          if (opts.verbose) process.stderr.write(`  ${label(job.input, "<stdin>")}: ${msg}
`);
        }
      });
      converted++;
      if (opts.verbose) {
        info(`  ${label(job.input, "<stdin>")}: done in ${((Date.now() - started) / 1e3).toFixed(1)}s`);
      }
    } catch (err) {
      if (controller.signal.aborted) return;
      fail(`${describe(job)}: ${err instanceof Error ? err.message : String(err)}`);
    }
  };
  let next = 0;
  const worker = async () => {
    while (next < jobs.length && !stopped && !controller.signal.aborted) {
      await processJob(jobs[next++]);
    }
  };
  await Promise.all(Array.from({ length: Math.min(opts.jobs, jobs.length) }, worker));
  if (controller.signal.aborted) {
    error("cancelled");
    return 130;
  }
  if (multi || opts.verbose) {
    const parts = [`${converted} converted`];
    if (skipped) parts.push(`${skipped} up to date`);
    if (failed) parts.push(`${failed} failed`);
    const notRun = jobs.length - converted - skipped - failed;
    if (notRun) parts.push(`${notRun} not attempted`);
    info(parts.join(", "));
  }
  return failed > 0 ? 1 : 0;
}
async function main() {
  let opts;
  try {
    opts = parseCliArgs(process.argv.slice(2));
  } catch (err) {
    if (!(err instanceof UsageError)) throw err;
    process.stderr.write(`error: ${err.message}
Try --help for usage.
`);
    process.exit(2);
  }
  if (opts.help) {
    process.stdout.write(USAGE);
    return;
  }
  if (opts.version) {
    process.stdout.write(
      `marp-to-editable-pptx-cli ${"0.1.0"} (marp-to-editable-pptx ${"1.4.1"})
`
    );
    return;
  }
  if (opts.inputs.length === 0 && process.stdin.isTTY) {
    process.stderr.write(USAGE);
    process.exit(2);
  }
  try {
    process.exitCode = await run(opts);
  } catch (err) {
    if (err instanceof UsageError) {
      process.stderr.write(`error: ${err.message}
`);
      process.exitCode = 2;
      return;
    }
    throw err;
  }
}
if (require.main === module) {
  main().catch((err) => {
    process.stderr.write(`error: ${err instanceof Error ? err.stack : String(err)}
`);
    process.exit(1);
  });
}
// Annotate the CommonJS export names for ESM import in node:
0 && (module.exports = {
  detectBrowserPath
});
