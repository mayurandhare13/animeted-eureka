/******/ (function() { // webpackBootstrap
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	!function() {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = function(exports) {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!********************************!*\
  !*** ./src/taskpane/dialog.ts ***!
  \********************************/
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { if (r) i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n;else { var o = function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); }; o("next", 0), o("throw", 1), o("return", 2); } }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/* global document, Office, console, navigator */

var requestedAction = "request";
Office.onReady(function () {
  var actionButton = document.getElementById("action-button");
  if (actionButton) {
    actionButton.onclick = function () {
      return runAction();
    };
  }
  Office.context.ui.addHandlerAsync(Office.EventType.DialogParentMessageReceived, function (arg) {
    try {
      var message = JSON.parse(arg.message);
      if (message.type === "start") {
        requestedAction = message.action === "devices" ? "devices" : "request";
        updateDialogLabel();
      }
    } catch (error) {
      console.error("Dialog message parse failed:", error);
    }
  });
  updateDialogLabel();
  Office.context.ui.messageParent(JSON.stringify({
    type: "ready"
  }));
});
function updateDialogLabel() {
  var label = document.getElementById("action-label");
  var button = document.getElementById("action-button");
  if (label) {
    label.textContent = requestedAction === "devices" ? "Enumerate microphone devices" : "Request microphone access";
  }
  if (button) {
    button.textContent = requestedAction === "devices" ? "Get devices" : "Request microphone";
  }
}
function runAction() {
  return _runAction.apply(this, arguments);
}
function _runAction() {
  _runAction = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
    var button, stream, devices, audioInputs, payload, message, denied, _t;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          button = document.getElementById("action-button");
          if (button) {
            button.disabled = true;
          }
          setStatus("Requesting microphone access...");
          _context.p = 1;
          _context.n = 2;
          return navigator.mediaDevices.getUserMedia({
            audio: true
          });
        case 2:
          stream = _context.v;
          if (!(requestedAction === "devices")) {
            _context.n = 4;
            break;
          }
          _context.n = 3;
          return navigator.mediaDevices.enumerateDevices();
        case 3:
          devices = _context.v;
          audioInputs = devices.filter(function (device) {
            return device.kind === "audioinput";
          });
          stream.getTracks().forEach(function (track) {
            return track.stop();
          });
          payload = audioInputs.map(function (device) {
            return {
              label: device.label || "(empty label)",
              deviceId: device.deviceId
            };
          });
          Office.context.ui.messageParent(JSON.stringify({
            type: "devices",
            devices: payload
          }));
          _context.n = 5;
          break;
        case 4:
          stream.getTracks().forEach(function (track) {
            return track.stop();
          });
          Office.context.ui.messageParent(JSON.stringify({
            type: "permission-granted"
          }));
        case 5:
          _context.n = 7;
          break;
        case 6:
          _context.p = 6;
          _t = _context.v;
          message = _t instanceof Error ? _t.message : String(_t);
          denied = message.includes("NotAllowedError") || message.toLowerCase().includes("denied");
          Office.context.ui.messageParent(JSON.stringify({
            type: denied ? "permission-denied" : "error",
            error: message
          }));
        case 7:
          _context.p = 7;
          if (button) {
            button.disabled = false;
          }
          return _context.f(7);
        case 8:
          return _context.a(2);
      }
    }, _callee, null, [[1, 6, 7, 8]]);
  }));
  return _runAction.apply(this, arguments);
}
function setStatus(message) {
  var status = document.getElementById("status");
  if (status) {
    status.textContent = message;
  }
}
}();
// This entry needs to be wrapped in an IIFE because it needs to be in strict mode.
!function() {
"use strict";
/*!**********************************!*\
  !*** ./src/taskpane/dialog.html ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\r\n<!DOCTYPE html>\r\n<html>\r\n\r\n<head>\r\n    <meta charset=\"UTF-8\" />\r\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\r\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\" />\r\n    <title>Microphone Dialog</title>\r\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\r\n    <style>\r\n        html, body {\r\n            margin: 0;\r\n            padding: 0;\r\n            font-family: Segoe UI, Arial, sans-serif;\r\n            color: #323130;\r\n        }\r\n\r\n        .container {\r\n            padding: 16px;\r\n        }\r\n\r\n        .title {\r\n            font-size: 16px;\r\n            font-weight: 600;\r\n            margin-bottom: 8px;\r\n        }\r\n\r\n        .subtitle {\r\n            font-size: 12px;\r\n            color: #605e5c;\r\n            margin-bottom: 12px;\r\n        }\r\n\r\n        .status {\r\n            margin-top: 10px;\r\n            font-size: 12px;\r\n        }\r\n\r\n        .btn {\r\n            display: inline-block;\r\n            padding: 10px 14px;\r\n            border-radius: 4px;\r\n            border: 1px solid #0078d4;\r\n            background: #0078d4;\r\n            color: #fff;\r\n            cursor: pointer;\r\n        }\r\n\r\n        .btn:disabled {\r\n            opacity: 0.6;\r\n            cursor: not-allowed;\r\n        }\r\n    </style>\r\n</head>\r\n\r\n<body>\r\n    <div class=\"container\">\r\n        <div class=\"title\" id=\"action-label\">Microphone access</div>\r\n        <div class=\"subtitle\">This dialog runs in a top-level window to request browser permission.</div>\r\n        <button id=\"action-button\" class=\"btn\" type=\"button\">Request microphone</button>\r\n        <div id=\"status\" class=\"status\"></div>\r\n    </div>\r\n</body>\r\n\r\n</html>\r\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=dialog.js.map