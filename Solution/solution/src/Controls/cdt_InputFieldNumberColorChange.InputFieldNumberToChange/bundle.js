/*
 * ATTENTION: The "eval" devtool has been used (maybe by default in mode: "development").
 * This devtool is neither made for production nor for readable output files.
 * It uses "eval()" calls to create a separate source file in the browser devtools.
 * If you are trying to read the output file, select a different devtool (https://webpack.js.org/configuration/devtool/)
 * or disable the default devtool with "devtool: false".
 * If you are looking for production-ready output files, see mode: "production" (https://webpack.js.org/configuration/mode/).
 */
var pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad;
/******/ (() => { // webpackBootstrap
/******/ 	"use strict";
/******/ 	var __webpack_modules__ = ({

/***/ "./InputFieldToChange/index.ts":
/*!*************************************!*\
  !*** ./InputFieldToChange/index.ts ***!
  \*************************************/
/***/ ((__unused_webpack_module, exports) => {

eval("\n\nObject.defineProperty(exports, \"__esModule\", ({\n  value: true\n}));\nexports.InputFieldNumberToChange = void 0;\nvar InputFieldNumberToChange = /** @class */function () {\n  /**\r\n   * Empty constructor.\r\n   */\n  function InputFieldNumberToChange() {\n    var _this = this;\n    /**\r\n     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.\r\n     * Data-set values are not initialized here, use updateView.\r\n     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.\r\n     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.\r\n     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.\r\n     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.\r\n     */\n    this.onBlur = function (event) {\n      _this._inputValue = Number(_this._inputElement.value);\n      _this._notifyOutputChanged();\n    };\n  }\n  InputFieldNumberToChange.prototype.init = function (context, notifyOutputChanged, state, container) {\n    // Add control initialization code\n    this._inputElement = document.createElement(\"input\");\n    this._inputElement.setAttribute(\"type\", \"number\");\n    this._inputElement.setAttribute(\"placeholder\", \"\");\n    this._inputElement.setAttribute(\"class\", \"textInputControl\");\n    this._inputValue = Number(context.parameters.InputFieldNumberToChange.raw) || 0;\n    this._inputElement.value = String(this._inputValue);\n    this._notifyOutputChanged = notifyOutputChanged;\n    this._inputElement.addEventListener(\"blur\", this.onBlur);\n    container.appendChild(this._inputElement);\n    this._colorInRgb = context.parameters.ColorInRGB.raw || \"rgb(198, 239, 206)\";\n    this._inputElement.style.backgroundColor = this._colorInRgb;\n  };\n  /**\r\n   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.\r\n   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions\r\n   */\n  InputFieldNumberToChange.prototype.updateView = function (context) {\n    // Add code to update control view\n    this._inputValue = Number(context.parameters.InputFieldNumberToChange.raw) || 0;\n    this._inputElement.value = String(this._inputValue);\n    this._colorInRgb = context.parameters.ColorInRGB.raw || \"\";\n    this._inputElement.style.backgroundColor = this._colorInRgb;\n  };\n  /**\r\n   * It is called by the framework prior to a control receiving new data.\r\n   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”\r\n   */\n  InputFieldNumberToChange.prototype.getOutputs = function () {\n    return {\n      InputFieldNumberToChange: this._inputValue\n    };\n  };\n  /**\r\n   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.\r\n   * i.e. cancelling any pending remote calls, removing listeners, etc.\r\n   */\n  InputFieldNumberToChange.prototype.destroy = function () {\n    // Add code to cleanup control if necessary\n  };\n  return InputFieldNumberToChange;\n}();\nexports.InputFieldNumberToChange = InputFieldNumberToChange;\n\n//# sourceURL=webpack://pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad/./InputFieldToChange/index.ts?");

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module can't be inlined because the eval devtool is used.
/******/ 	var __webpack_exports__ = {};
/******/ 	__webpack_modules__["./InputFieldToChange/index.ts"](0, __webpack_exports__);
/******/ 	pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad = __webpack_exports__;
/******/ 	
/******/ })()
;
if (window.ComponentFramework && window.ComponentFramework.registerControl) {
	ComponentFramework.registerControl('InputFieldNumberColorChange.InputFieldNumberToChange', pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad.InputFieldNumberToChange);
} else {
	var InputFieldNumberColorChange = InputFieldNumberColorChange || {};
	InputFieldNumberColorChange.InputFieldNumberToChange = pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad.InputFieldNumberToChange;
	pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad = undefined;
}