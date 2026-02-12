/******/ (function() { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./src/taskpane/officethemes.css":
/*!***************************************!*\
  !*** ./src/taskpane/officethemes.css ***!
  \***************************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "officethemes.css";

/***/ }),

/***/ "./src/taskpane/taskpane.css":
/*!***********************************!*\
  !*** ./src/taskpane/taskpane.css ***!
  \***********************************/
/***/ (function(module, __unused_webpack_exports, __webpack_require__) {

module.exports = __webpack_require__.p + "taskpane.css";

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The module cache
/******/ 	var __webpack_module_cache__ = {};
/******/ 	
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/ 		// Check if module is in cache
/******/ 		var cachedModule = __webpack_module_cache__[moduleId];
/******/ 		if (cachedModule !== undefined) {
/******/ 			return cachedModule.exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = __webpack_module_cache__[moduleId] = {
/******/ 			// no module.id needed
/******/ 			// no module.loaded needed
/******/ 			exports: {}
/******/ 		};
/******/ 	
/******/ 		// Execute the module function
/******/ 		__webpack_modules__[moduleId](module, module.exports, __webpack_require__);
/******/ 	
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/ 	
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = __webpack_modules__;
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	!function() {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = function(exports, definition) {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/global */
/******/ 	!function() {
/******/ 		__webpack_require__.g = (function() {
/******/ 			if (typeof globalThis === 'object') return globalThis;
/******/ 			try {
/******/ 				return this || new Function('return this')();
/******/ 			} catch (e) {
/******/ 				if (typeof window === 'object') return window;
/******/ 			}
/******/ 		})();
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	!function() {
/******/ 		__webpack_require__.o = function(obj, prop) { return Object.prototype.hasOwnProperty.call(obj, prop); }
/******/ 	}();
/******/ 	
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
/******/ 	/* webpack/runtime/publicPath */
/******/ 	!function() {
/******/ 		var scriptUrl;
/******/ 		if (__webpack_require__.g.importScripts) scriptUrl = __webpack_require__.g.location + "";
/******/ 		var document = __webpack_require__.g.document;
/******/ 		if (!scriptUrl && document) {
/******/ 			if (document.currentScript && document.currentScript.tagName.toUpperCase() === 'SCRIPT')
/******/ 				scriptUrl = document.currentScript.src;
/******/ 			if (!scriptUrl) {
/******/ 				var scripts = document.getElementsByTagName("script");
/******/ 				if(scripts.length) {
/******/ 					var i = scripts.length - 1;
/******/ 					while (i > -1 && (!scriptUrl || !/^http(s?):/.test(scriptUrl))) scriptUrl = scripts[i--].src;
/******/ 				}
/******/ 			}
/******/ 		}
/******/ 		// When supporting browsers where an automatic publicPath is not supported you must specify an output.publicPath manually via configuration
/******/ 		// or pass an empty string ("") and set the __webpack_public_path__ variable from your code to use your own logic.
/******/ 		if (!scriptUrl) throw new Error("Automatic publicPath is not supported in this browser");
/******/ 		scriptUrl = scriptUrl.replace(/^blob:/, "").replace(/#.*$/, "").replace(/\?.*$/, "").replace(/\/[^\/]+$/, "/");
/******/ 		__webpack_require__.p = scriptUrl;
/******/ 	}();
/******/ 	
/******/ 	/* webpack/runtime/jsonp chunk loading */
/******/ 	!function() {
/******/ 		__webpack_require__.b = document.baseURI || self.location.href;
/******/ 		
/******/ 		// object to store loaded and loading chunks
/******/ 		// undefined = chunk not loaded, null = chunk preloaded/prefetched
/******/ 		// [resolve, reject, Promise] = chunk loading, 0 = chunk loaded
/******/ 		var installedChunks = {
/******/ 			"taskpane": 0
/******/ 		};
/******/ 		
/******/ 		// no chunk on demand loading
/******/ 		
/******/ 		// no prefetching
/******/ 		
/******/ 		// no preloaded
/******/ 		
/******/ 		// no HMR
/******/ 		
/******/ 		// no HMR manifest
/******/ 		
/******/ 		// no on chunks loaded
/******/ 		
/******/ 		// no jsonp function
/******/ 	}();
/******/ 	
/************************************************************************/
var __webpack_exports__ = {};
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
var __webpack_exports__ = {};
/*!**********************************!*\
  !*** ./src/taskpane/taskpane.ts ***!
  \**********************************/
__webpack_require__.r(__webpack_exports__);
/* harmony export */ __webpack_require__.d(__webpack_exports__, {
/* harmony export */   checkPermissionStatus: function() { return /* binding */ checkPermissionStatus; },
/* harmony export */   getMicrophoneDevices: function() { return /* binding */ getMicrophoneDevices; },
/* harmony export */   requestMicrophonePermission: function() { return /* binding */ requestMicrophonePermission; }
/* harmony export */ });
function _regenerator() { /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/babel/babel/blob/main/packages/babel-helpers/LICENSE */ var e, t, r = "function" == typeof Symbol ? Symbol : {}, n = r.iterator || "@@iterator", o = r.toStringTag || "@@toStringTag"; function i(r, n, o, i) { var c = n && n.prototype instanceof Generator ? n : Generator, u = Object.create(c.prototype); return _regeneratorDefine2(u, "_invoke", function (r, n, o) { var i, c, u, f = 0, p = o || [], y = !1, G = { p: 0, n: 0, v: e, a: d, f: d.bind(e, 4), d: function d(t, r) { return i = t, c = 0, u = e, G.n = r, a; } }; function d(r, n) { for (c = r, u = n, t = 0; !y && f && !o && t < p.length; t++) { var o, i = p[t], d = G.p, l = i[2]; r > 3 ? (o = l === n) && (u = i[(c = i[4]) ? 5 : (c = 3, 3)], i[4] = i[5] = e) : i[0] <= d && ((o = r < 2 && d < i[1]) ? (c = 0, G.v = n, G.n = i[1]) : d < l && (o = r < 3 || i[0] > n || n > l) && (i[4] = r, i[5] = n, G.n = l, c = 0)); } if (o || r > 1) return a; throw y = !0, n; } return function (o, p, l) { if (f > 1) throw TypeError("Generator is already running"); for (y && 1 === p && d(p, l), c = p, u = l; (t = c < 2 ? e : u) || !y;) { i || (c ? c < 3 ? (c > 1 && (G.n = -1), d(c, u)) : G.n = u : G.v = u); try { if (f = 2, i) { if (c || (o = "next"), t = i[o]) { if (!(t = t.call(i, u))) throw TypeError("iterator result is not an object"); if (!t.done) return t; u = t.value, c < 2 && (c = 0); } else 1 === c && (t = i.return) && t.call(i), c < 2 && (u = TypeError("The iterator does not provide a '" + o + "' method"), c = 1); i = e; } else if ((t = (y = G.n < 0) ? u : r.call(n, G)) !== a) break; } catch (t) { i = e, c = 1, u = t; } finally { f = 1; } } return { value: t, done: y }; }; }(r, o, i), !0), u; } var a = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} t = Object.getPrototypeOf; var c = [][n] ? t(t([][n]())) : (_regeneratorDefine2(t = {}, n, function () { return this; }), t), u = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(c); function f(e) { return Object.setPrototypeOf ? Object.setPrototypeOf(e, GeneratorFunctionPrototype) : (e.__proto__ = GeneratorFunctionPrototype, _regeneratorDefine2(e, o, "GeneratorFunction")), e.prototype = Object.create(u), e; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, _regeneratorDefine2(u, "constructor", GeneratorFunctionPrototype), _regeneratorDefine2(GeneratorFunctionPrototype, "constructor", GeneratorFunction), GeneratorFunction.displayName = "GeneratorFunction", _regeneratorDefine2(GeneratorFunctionPrototype, o, "GeneratorFunction"), _regeneratorDefine2(u), _regeneratorDefine2(u, o, "Generator"), _regeneratorDefine2(u, n, function () { return this; }), _regeneratorDefine2(u, "toString", function () { return "[object Generator]"; }), (_regenerator = function _regenerator() { return { w: i, m: f }; })(); }
function _regeneratorDefine2(e, r, n, t) { var i = Object.defineProperty; try { i({}, "", {}); } catch (e) { i = 0; } _regeneratorDefine2 = function _regeneratorDefine(e, r, n, t) { if (r) i ? i(e, r, { value: n, enumerable: !t, configurable: !t, writable: !t }) : e[r] = n;else { var o = function o(r, n) { _regeneratorDefine2(e, r, function (e) { return this._invoke(r, n, e); }); }; o("next", 0), o("throw", 1), o("return", 2); } }, _regeneratorDefine2(e, r, n, t); }
function asyncGeneratorStep(n, t, e, r, o, a, c) { try { var i = n[a](c), u = i.value; } catch (n) { return void e(n); } i.done ? t(u) : Promise.resolve(u).then(r, o); }
function _asyncToGenerator(n) { return function () { var t = this, e = arguments; return new Promise(function (r, o) { var a = n.apply(t, e); function _next(n) { asyncGeneratorStep(a, r, o, _next, _throw, "next", n); } function _throw(n) { asyncGeneratorStep(a, r, o, _next, _throw, "throw", n); } _next(void 0); }); }; }
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, console */

Office.onReady(function (info) {
  if (info.host === Office.HostType.PowerPoint) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Add event listeners for simplified API buttons
    document.getElementById("check-permission").onclick = checkPermissionStatus;
    document.getElementById("request-permission").onclick = requestMicrophonePermission;
    document.getElementById("get-devices").onclick = getMicrophoneDevices;
  }
});
var hasShownPermissionHint = false;
// ============================================
// 1. Check Permission Status
// ============================================
function checkPermissionStatus() {
  return _checkPermissionStatus.apply(this, arguments);
}

// ============================================
// 2. Request Microphone Permission
// ============================================
function _checkPermissionStatus() {
  _checkPermissionStatus = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee() {
    var permissionStatus, _t, _t2;
    return _regenerator().w(function (_context) {
      while (1) switch (_context.n) {
        case 0:
          _context.p = 0;
          console.log("Checking permission status...");
          updateStatus("Checking microphone permission status...", "info");

          // Check if API is available
          if (Office.devicePermission) {
            _context.n = 1;
            break;
          }
          updateStatus("❌ Device Permission API is not available", "error");
          updatePermissionType("Requires Office on the web (Excel, PowerPoint, Word, or Outlook)");
          return _context.a(2);
        case 1:
          if (!(Office.context.platform !== Office.PlatformType.OfficeOnline)) {
            _context.n = 2;
            break;
          }
          updateStatus("❌ Not running on Office on the web", "error");
          updatePermissionType("Current platform: " + Office.context.platform);
          return _context.a(2);
        case 2:
          if (!navigator.permissions) {
            _context.n = 7;
            break;
          }
          _context.p = 3;
          _context.n = 4;
          return navigator.permissions.query({
            name: 'microphone'
          });
        case 4:
          permissionStatus = _context.v;
          console.log("Permission state:", permissionStatus.state);
          if (permissionStatus.state === 'granted') {
            updateStatus("✅ Microphone permission is GRANTED", "success");
            updatePermissionType("Permission is active. Type: Cannot determine if 'Always' or 'Allow Once' from browser API alone.");
            hidePermissionHint();
          } else if (permissionStatus.state === 'denied') {
            updateStatus("⛔ Microphone permission is DENIED", "error");
            updatePermissionType("User has blocked microphone access. Clear site permissions in browser settings.");
            hidePermissionHint();
          } else {
            updateStatus("⚠️ Microphone permission is PROMPT (not yet decided)", "info");
            updatePermissionType("User hasn't granted or denied permission yet. Click 'Request Permission' to prompt.");
            showPermissionHintOnce("Tip: Click 'Request Microphone Permission' to trigger the browser prompt before trying to list devices.");
          }

          // Listen for permission changes
          permissionStatus.onchange = function () {
            console.log("Permission changed to:", permissionStatus.state);
            checkPermissionStatus(); // Recheck when permission changes
          };
          _context.n = 6;
          break;
        case 5:
          _context.p = 5;
          _t = _context.v;
          console.error("Could not query permissions:", _t);
          updateStatus("✅ API is available, but cannot check permission state", "info");
          updatePermissionType("Click 'Request Permission' to proceed.");
        case 6:
          _context.n = 8;
          break;
        case 7:
          updateStatus("✅ API is available", "success");
          updatePermissionType("navigator.permissions API not available. Click 'Request Permission' to proceed.");
        case 8:
          _context.n = 10;
          break;
        case 9:
          _context.p = 9;
          _t2 = _context.v;
          console.error("Check permission error:", _t2);
          updateStatus("Error checking permission: ".concat(_t2 instanceof Error ? _t2.message : _t2), "error");
        case 10:
          return _context.a(2);
      }
    }, _callee, null, [[3, 5], [0, 9]]);
  }));
  return _checkPermissionStatus.apply(this, arguments);
}
function requestMicrophonePermission() {
  return _requestMicrophonePermission.apply(this, arguments);
}

// ============================================
// 3. Get Microphone Devices
// ============================================
function _requestMicrophonePermission() {
  _requestMicrophonePermission = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee2() {
    var deviceCapabilities, isFirstTimeGrant, dialogResult, errorMessage, _t3;
    return _regenerator().w(function (_context2) {
      while (1) switch (_context2.n) {
        case 0:
          _context2.p = 0;
          console.log("Requesting microphone permission...");
          updateStatus("Requesting microphone permission...", "info");
          updatePermissionType("");

          // Check if API is available
          if (Office.devicePermission) {
            _context2.n = 1;
            break;
          }
          updateStatus("❌ Device Permission API is not available", "error");
          updatePermissionType("Requires Office on the web");
          return _context2.a(2);
        case 1:
          if (!(Office.context.platform !== Office.PlatformType.OfficeOnline)) {
            _context2.n = 2;
            break;
          }
          updateStatus("❌ Only supported in Office on the web", "error");
          updatePermissionType("Current platform: " + Office.context.platform);
          return _context2.a(2);
        case 2:
          // Request permission
          deviceCapabilities = [Office.DevicePermissionType.microphone];
          _context2.n = 3;
          return Office.devicePermission.requestPermissions(deviceCapabilities);
        case 3:
          isFirstTimeGrant = _context2.v;
          if (!isFirstTimeGrant) {
            _context2.n = 4;
            break;
          }
          // User clicked "Allow" or "Allow Once" for the first time
          updateStatus("✅ Permission granted! Reloading in 2 seconds...", "success");
          updatePermissionType("First-time grant detected. The add-in will reload to apply permissions.");
          hidePermissionHint();
          setTimeout(function () {
            console.log("Reloading after first-time permission grant...");
            location.reload();
          }, 2000);
          _context2.n = 8;
          break;
        case 4:
          // Permission was already granted previously
          updateStatus("✅ Permission already granted (using existing permission)", "success");
          updatePermissionType("This permission was granted earlier. Type: 'Always' if persistent after tab close, 'Allow Once' if requires re-grant after tab close.");
          hidePermissionHint();

          // Test that we can access microphone
          _context2.n = 5;
          return waitForPermissionPropagation();
        case 5:
          if (!useDialogFlow()) {
            _context2.n = 7;
            break;
          }
          _context2.n = 6;
          return openMicDialog("request");
        case 6:
          dialogResult = _context2.v;
          handleDialogPermissionResult(dialogResult);
          return _context2.a(2);
        case 7:
          // if (isSafari()) {
          //   updateStatus("⚠️ Safari blocks microphone access in Office on the web", "error");
          //   updatePermissionType(
          //     "This is a permissions-policy restriction from the host. Use Chrome/Firefox or desktop to test getUserMedia."
          //   );
          //   return;
          // }

          testMicrophoneAccess();
        case 8:
          _context2.n = 10;
          break;
        case 9:
          _context2.p = 9;
          _t3 = _context2.v;
          console.error("Request permission error:", _t3);
          errorMessage = _t3 instanceof Error ? _t3.message : String(_t3);
          if (errorMessage.includes("denied")) {
            updateStatus("⛔ Permission DENIED by user", "error");
            updatePermissionType("User clicked 'Deny'. You can request permission again, and the user will be prompted again.");
          } else {
            updateStatus("\u274C Error: ".concat(errorMessage), "error");
            updatePermissionType("");
          }
        case 10:
          return _context2.a(2);
      }
    }, _callee2, null, [[0, 9]]);
  }));
  return _requestMicrophonePermission.apply(this, arguments);
}
function getMicrophoneDevices() {
  return _getMicrophoneDevices.apply(this, arguments);
}

// ============================================
// Helper Functions
// ============================================
function _getMicrophoneDevices() {
  _getMicrophoneDevices = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee3() {
    var dialogResult, stream, devices, audioInputs, errorMsg, _t4;
    return _regenerator().w(function (_context3) {
      while (1) switch (_context3.n) {
        case 0:
          _context3.p = 0;
          console.log("Getting microphone devices...");
          updateStatus("Enumerating microphone devices...", "info");
          if (!useDialogFlow()) {
            _context3.n = 2;
            break;
          }
          _context3.n = 1;
          return openMicDialog("devices");
        case 1:
          dialogResult = _context3.v;
          handleDialogDevicesResult(dialogResult);
          return _context3.a(2);
        case 2:
          if (navigator.mediaDevices) {
            _context3.n = 3;
            break;
          }
          updateStatus("❌ MediaDevices API not available", "error");
          return _context3.a(2);
        case 3:
          _context3.n = 4;
          return navigator.mediaDevices.getUserMedia({
            audio: true
          });
        case 4:
          stream = _context3.v;
          console.log("✅ Microphone access granted");

          // Enumerate devices while stream is active - labels will be populated
          _context3.n = 5;
          return navigator.mediaDevices.enumerateDevices();
        case 5:
          devices = _context3.v;
          audioInputs = devices.filter(function (device) {
            return device.kind === "audioinput";
          }); // Stop the stream immediately after enumeration
          stream.getTracks().forEach(function (track) {
            return track.stop();
          });
          console.log("Found ".concat(audioInputs.length, " audio input device(s)"));
          if (!(audioInputs.length === 0)) {
            _context3.n = 6;
            break;
          }
          updateStatus("⚠️ No microphone devices found", "error");
          updatePermissionType("Make sure a microphone is connected.");
          clearDeviceList();
          return _context3.a(2);
        case 6:
          updateStatus("\u2705 Found ".concat(audioInputs.length, " microphone device(s)"), "success");
          updatePermissionType("Devices are accessible with permission granted.");
          displayDevices(audioInputs.map(function (device) {
            return {
              label: device.label || "(empty label)",
              deviceId: device.deviceId
            };
          }), true);
          _context3.n = 8;
          break;
        case 7:
          _context3.p = 7;
          _t4 = _context3.v;
          console.error("Get devices error:", _t4);
          errorMsg = _t4 instanceof Error ? _t4.message : String(_t4);
          if (errorMsg.includes("Permission denied") || errorMsg.includes("NotAllowedError")) {
            updateStatus("⚠️ Microphone access denied. Click 'Request Permission' first.", "error");
            updatePermissionType("Permission not granted or user declined.");
          } else {
            updateStatus("\u274C Error: ".concat(errorMsg), "error");
          }
          clearDeviceList();
        case 8:
          return _context3.a(2);
      }
    }, _callee3, null, [[0, 7]]);
  }));
  return _getMicrophoneDevices.apply(this, arguments);
}
function updateStatus(message, type) {
  var statusElement = document.getElementById("permission-status");
  var statusText = document.getElementById("status-text");
  if (statusElement && statusText) {
    statusText.textContent = message;
    statusElement.style.display = "block";

    // Update styling based on type
    if (type === "success") {
      statusElement.style.borderLeft = "4px solid #107C10";
      statusElement.style.background = "#DFF6DD";
    } else if (type === "error") {
      statusElement.style.borderLeft = "4px solid #D13438";
      statusElement.style.background = "#FDE7E9";
    } else {
      statusElement.style.borderLeft = "4px solid #0078D4";
      statusElement.style.background = "#F3F2F1";
    }
  }
  console.log("[".concat(type.toUpperCase(), "] ").concat(message));
}
function updatePermissionType(message) {
  var permissionTypeElement = document.getElementById("permission-type");
  if (permissionTypeElement) {
    permissionTypeElement.textContent = message;
    permissionTypeElement.style.display = message ? "block" : "none";
  }
}
function displayDevices(devices, hasLabels) {
  var deviceListContainer = document.getElementById("device-list");
  var devicesList = document.getElementById("devices");
  if (deviceListContainer && devicesList) {
    devicesList.innerHTML = "";
    devices.forEach(function (device, index) {
      var li = document.createElement("li");
      li.style.padding = "8px";
      li.style.marginBottom = "4px";
      li.style.background = "#F3F2F1";
      li.style.borderRadius = "4px";
      var label = device.label || "(empty label)";
      var deviceId = device.deviceId.substring(0, 20) + "...";
      li.innerHTML = "\n        <div style=\"font-weight: 600;\">".concat(index + 1, ". ").concat(hasLabels ? "🎤" : "⚠️", " ").concat(label, "</div>\n        <div style=\"font-size: 11px; color: #605E5C; margin-top: 4px;\">ID: ").concat(deviceId, "</div>\n      ");
      devicesList.appendChild(li);
    });
    deviceListContainer.style.display = "block";
  }
}
function clearDeviceList() {
  var deviceListContainer = document.getElementById("device-list");
  if (deviceListContainer) {
    deviceListContainer.style.display = "none";
  }
}
function showPermissionHintOnce(message) {
  if (hasShownPermissionHint) {
    return;
  }
  var hintElement = document.getElementById("permission-hint");
  if (hintElement) {
    hintElement.textContent = message;
    hintElement.style.display = "block";
    hasShownPermissionHint = true;
  }
}
function hidePermissionHint() {
  var hintElement = document.getElementById("permission-hint");
  if (hintElement) {
    hintElement.style.display = "none";
  }
}
function testMicrophoneAccess() {
  return _testMicrophoneAccess.apply(this, arguments);
}
function _testMicrophoneAccess() {
  _testMicrophoneAccess = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee4() {
    var stream, _t5;
    return _regenerator().w(function (_context4) {
      while (1) switch (_context4.n) {
        case 0:
          _context4.p = 0;
          console.log("Testing microphone access...");
          // if (isSafari()) {
          //   console.warn("Safari blocks getUserMedia in Office on the web due to permissions-policy.");
          //   return;
          // }
          // await waitForPermissionPropagation();
          _context4.n = 1;
          return navigator.mediaDevices.getUserMedia({
            audio: true
          });
        case 1:
          stream = _context4.v;
          console.log("✅ Microphone access successful:", stream);

          // Stop the stream immediately (we just wanted to test access)
          stream.getTracks().forEach(function (track) {
            return track.stop();
          });
          _context4.n = 3;
          break;
        case 2:
          _context4.p = 2;
          _t5 = _context4.v;
          console.error("❌ Microphone access test failed:", _t5);
        case 3:
          return _context4.a(2);
      }
    }, _callee4, null, [[0, 2]]);
  }));
  return _testMicrophoneAccess.apply(this, arguments);
}
function isFirefox() {
  return navigator.userAgent.toLowerCase().includes("firefox");
}
function isSafari() {
  var ua = navigator.userAgent.toLowerCase();
  return ua.includes("safari") && !ua.includes("chrome") && !ua.includes("chromium") && !ua.includes("crios") && !ua.includes("fxios") && !ua.includes("edgios");
}
function useDialogFlow() {
  var toggle = document.getElementById("use-dialog-toggle");
  return Boolean(toggle && toggle.checked);
}
function getDialogUrl() {
  return new URL("dialog.html", window.location.href).toString();
}
function openMicDialog(action) {
  return new Promise(function (resolve) {
    Office.context.ui.displayDialogAsync(getDialogUrl(), {
      height: 45,
      width: 40,
      displayInIframe: false
    }, function (asyncResult) {
      if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        resolve({
          type: "error",
          error: asyncResult.error.message
        });
        return;
      }
      var dialog = asyncResult.value;
      var closeAndResolve = function closeAndResolve(result) {
        dialog.close();
        resolve(result);
      };
      dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
        try {
          var message = JSON.parse(arg.message);
          if (message.type === "ready") {
            dialog.messageChild(JSON.stringify({
              type: "start",
              action: action
            }));
            return;
          }
          closeAndResolve(message);
        } catch (error) {
          closeAndResolve({
            type: "error",
            error: "Failed to parse dialog response."
          });
        }
      });
      dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
        closeAndResolve({
          type: "error",
          error: "Dialog closed before completing the request."
        });
      });
    });
  });
}
function handleDialogPermissionResult(result) {
  if (result.type === "permission-granted") {
    updateStatus("✅ Microphone permission granted via dialog", "success");
    updatePermissionType("Top-level browser permission granted.");
    return;
  }
  if (result.type === "permission-denied") {
    updateStatus("⛔ Microphone permission denied in dialog", "error");
    updatePermissionType(result.error || "User denied permission.");
    return;
  }
  if (result.type === "error") {
    updateStatus("\u274C Dialog error: ".concat(result.error), "error");
    updatePermissionType("");
  }
}
function handleDialogDevicesResult(result) {
  if (result.type === "devices") {
    updateStatus("\u2705 Found ".concat(result.devices.length, " microphone device(s)"), "success");
    updatePermissionType("Devices are accessible with dialog permission.");
    displayDevices(result.devices, true);
    return;
  }
  if (result.type === "permission-denied") {
    updateStatus("⚠️ Microphone access denied in dialog.", "error");
    updatePermissionType(result.error || "User denied permission.");
    clearDeviceList();
    return;
  }
  if (result.type === "error") {
    updateStatus("\u274C Dialog error: ".concat(result.error), "error");
    updatePermissionType("");
    clearDeviceList();
  }
}
function waitForPermissionPropagation() {
  return _waitForPermissionPropagation.apply(this, arguments);
}
function _waitForPermissionPropagation() {
  _waitForPermissionPropagation = _asyncToGenerator(/*#__PURE__*/_regenerator().m(function _callee5() {
    return _regenerator().w(function (_context5) {
      while (1) switch (_context5.n) {
        case 0:
          if (isFirefox()) {
            _context5.n = 1;
            break;
          }
          return _context5.a(2);
        case 1:
          _context5.n = 2;
          return new Promise(function (resolve) {
            requestAnimationFrame(function () {
              return resolve();
            });
          });
        case 2:
          return _context5.a(2);
      }
    }, _callee5);
  }));
  return _waitForPermissionPropagation.apply(this, arguments);
}
}();
// This entry needs to be wrapped in an IIFE because it needs to be isolated against other entry modules.
!function() {
/*!************************************!*\
  !*** ./src/taskpane/taskpane.html ***!
  \************************************/
__webpack_require__.r(__webpack_exports__);
// Imports
var ___HTML_LOADER_IMPORT_0___ = new URL(/* asset import */ __webpack_require__(/*! ./officethemes.css */ "./src/taskpane/officethemes.css"), __webpack_require__.b);
var ___HTML_LOADER_IMPORT_1___ = new URL(/* asset import */ __webpack_require__(/*! ./taskpane.css */ "./src/taskpane/taskpane.css"), __webpack_require__.b);
// Module
var code = "<!-- Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT License. -->\n<!DOCTYPE html>\n<html>\n\n<head>\n    <meta charset=\"UTF-8\" />\n    <meta http-equiv=\"X-UA-Compatible\" content=\"IE=Edge\" />\n    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n    <title>Device Permission Repro</title>\n    <link href=\"" + ___HTML_LOADER_IMPORT_0___ + "\" rel=\"stylesheet\" type=\"text/css\" />\n    <" + "script type=\"text/javascript\" src=\"https://appsforoffice.microsoft.com/lib/1/hosted/office.js\"><" + "/script>\n    <link rel=\"stylesheet\" href=\"https://res-1.cdn.office.net/files/fabric-cdn-prod_20230815.002/office-ui-fabric-core/11.1.0/css/fabric.min.css\"/>\n    <link href=\"" + ___HTML_LOADER_IMPORT_1___ + "\" rel=\"stylesheet\" type=\"text/css\" />\n    <style>\n        .btn-group { display: flex; flex-direction: column; gap: 8px; margin: 12px 0; }\n        .btn-group .ms-Button { width: 100%; }\n        .step-label { font-weight: 600; margin-right: 6px; }\n        #status { max-height: 300px; overflow-y: auto; font-family: monospace; font-size: 12px; margin-top: 12px; }\n        .status-success { color: green; }\n        .status-error { color: red; }\n        .status-info { color: #0078d4; }\n    </style>\n</head>\n\n<body class=\"ms-font-m ms-welcome ms-Fabric\">\n    <header class=\"ms-welcome__header ms-bgColor-neutralLighter\">\n        <h1 class=\"ms-font-su\">Mic Permission Repro</h1>\n        <p class=\"ms-font-m\">Issue <a href=\"https://github.com/OfficeDev/office-js/issues/5726\" target=\"_blank\">#5726</a></p>\n    </header>\n    <section id=\"sideload-msg\" class=\"ms-welcome__main\">\n        <h2 class=\"ms-font-xl\">Please sideload your add-in to see app body.</h2>\n    </section>\n    <main id=\"app-body\" class=\"ms-welcome__main\" style=\"display: none;\">\n        <h2 class=\"ms-font-xl\">Microphone Permission API</h2>\n        <p class=\"ms-font-m\">Simple interface to test Office DevicePermission API</p>\n\n        <div class=\"toggle-row\">\n            <input type=\"checkbox\" id=\"use-dialog-toggle\" />\n            <label for=\"use-dialog-toggle\">Use dialog (top-level) for microphone access</label>\n        </div>\n\n        <div id=\"permission-hint\" class=\"permission-hint\" style=\"display: none;\"></div>\n\n        <div class=\"btn-group\">\n            <div role=\"button\" id=\"check-permission\" class=\"ms-Button ms-Button--primary ms-font-m\">\n                <span class=\"ms-Button-label\">1. Check Permission Status</span>\n            </div>\n            <div role=\"button\" id=\"request-permission\" class=\"ms-Button ms-Button--hero ms-font-m\">\n                <span class=\"ms-Button-label\">2. Request Microphone Permission</span>\n            </div>\n            <div role=\"button\" id=\"get-devices\" class=\"ms-Button ms-Button--primary ms-font-m\">\n                <span class=\"ms-Button-label\">3. Get Microphone Devices</span>\n            </div>\n        </div>\n\n        <div id=\"permission-status\" style=\"display: none; margin: 12px 0; padding: 12px; border-radius: 4px; background: #f3f2f1;\">\n            <div id=\"status-text\" class=\"ms-font-m\"></div>\n            <div id=\"permission-type\" class=\"ms-font-s\" style=\"margin-top: 8px; font-style: italic;\"></div>\n        </div>\n\n        <div id=\"device-list\" style=\"display: none; margin: 12px 0;\">\n            <h3 class=\"ms-font-l\">Microphone Devices:</h3>\n            <ul id=\"devices\" style=\"list-style: none; padding: 0;\"></ul>\n        </div>\n    </main>\n</body>\n\n</html>\n";
// Exports
/* harmony default export */ __webpack_exports__["default"] = (code);
}();
/******/ })()
;
//# sourceMappingURL=taskpane.js.map