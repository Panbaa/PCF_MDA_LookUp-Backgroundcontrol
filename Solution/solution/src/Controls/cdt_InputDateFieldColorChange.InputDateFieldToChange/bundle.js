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

/***/ "./InputDateFieldColorChange/index.ts":
/*!********************************************!*\
  !*** ./InputDateFieldColorChange/index.ts ***!
  \********************************************/
/***/ ((__unused_webpack_module, __webpack_exports__, __webpack_require__) => {

eval("__webpack_require__.r(__webpack_exports__);\n/* harmony export */ __webpack_require__.d(__webpack_exports__, {\n/* harmony export */   InputDateFieldToChange: () => (/* binding */ InputDateFieldToChange)\n/* harmony export */ });\nclass InputDateFieldToChange {\n  /**\r\n   * Empty constructor.\r\n   */\n  constructor() {\n    /**\r\n     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.\r\n     * Data-set values are not initialized here, use updateView.\r\n     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.\r\n     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.\r\n     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.\r\n     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.\r\n     */\n    this.onChange = event => {\n      console.log(\"onChange\");\n      var twoHoursToAdd = 1000 * 60 * 60 * 2;\n      var dateFormatter = new Date(this._inputElement.value).getTime();\n      dateFormatter = dateFormatter + twoHoursToAdd;\n      this._inputValue = new Date(dateFormatter);\n      this._notifyOutputChanged();\n    };\n  }\n  init(context, notifyOutputChanged, state, container) {\n    // Add control initialization code\n    this._notifyOutputChanged = notifyOutputChanged;\n    console.log(\"init\");\n    // Create DOM Element\n    this._inputElement = document.createElement(\"input\");\n    this._inputElement.setAttribute(\"type\", \"date\");\n    this._inputElement.setAttribute(\"class\", \"textInputControl\");\n    container.appendChild(this._inputElement);\n    // Listen for blur event\n    this._inputElement.addEventListener(\"blur\", this.onChange);\n    // Initialize input value\n    var rawDateValue = context.parameters.InputDateFieldToChange.raw;\n    this._inputValue = rawDateValue ? new Date(rawDateValue) : null;\n    // Set input value\n    this._inputElement.value = this._inputValue ? String(this._inputValue) : \"\";\n    this._colorInRgb = context.parameters.ColorInRGB.raw || \"rgb(198, 239, 206)\";\n    this._inputElement.style.backgroundColor = this._colorInRgb;\n  }\n  /**\r\n   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.\r\n   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions\r\n   */\n  updateView(context) {\n    // Add code to update control view\n    console.log(\"update\");\n    this._inputValue = context.parameters.InputDateFieldToChange.raw || null; // new Date(0);\n    // Get year, month, and day part from the date\n    var year = this._inputValue ? this._inputValue.toLocaleString(\"default\", {\n      year: \"numeric\"\n    }) : \"\";\n    var month = this._inputValue ? this._inputValue.toLocaleString(\"default\", {\n      month: \"2-digit\"\n    }) : \"\";\n    var day = this._inputValue ? this._inputValue.toLocaleString(\"default\", {\n      day: \"2-digit\"\n    }) : \"\";\n    // Generate yyyy-mm-dd date string\n    var formattedDate = year + \"-\" + month + \"-\" + day;\n    this._inputElement.value = formattedDate;\n    this._colorInRgb = context.parameters.ColorInRGB.raw || \"\";\n    this._inputElement.style.backgroundColor = this._colorInRgb;\n  }\n  /**\r\n   * It is called by the framework prior to a control receiving new data.\r\n   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”\r\n   */\n  getOutputs() {\n    console.log(\"outputs\");\n    return {\n      InputDateFieldToChange: this._inputValue !== null ? new Date(this._inputValue) : undefined\n      // InputDateFieldToChange: new Date(this._inputValue),\n    };\n  }\n  /**\r\n   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.\r\n   * i.e. cancelling any pending remote calls, removing listeners, etc.\r\n   */\n  destroy() {\n    // Add code to cleanup control if necessary\n  }\n}\n\n//# sourceURL=webpack://pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad/./InputDateFieldColorChange/index.ts?");

/***/ })

/******/ 	});
/************************************************************************/
/******/ 	// The require scope
/******/ 	var __webpack_require__ = {};
/******/ 	
/************************************************************************/
/******/ 	/* webpack/runtime/define property getters */
/******/ 	(() => {
/******/ 		// define getter functions for harmony exports
/******/ 		__webpack_require__.d = (exports, definition) => {
/******/ 			for(var key in definition) {
/******/ 				if(__webpack_require__.o(definition, key) && !__webpack_require__.o(exports, key)) {
/******/ 					Object.defineProperty(exports, key, { enumerable: true, get: definition[key] });
/******/ 				}
/******/ 			}
/******/ 		};
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/hasOwnProperty shorthand */
/******/ 	(() => {
/******/ 		__webpack_require__.o = (obj, prop) => (Object.prototype.hasOwnProperty.call(obj, prop))
/******/ 	})();
/******/ 	
/******/ 	/* webpack/runtime/make namespace object */
/******/ 	(() => {
/******/ 		// define __esModule on exports
/******/ 		__webpack_require__.r = (exports) => {
/******/ 			if(typeof Symbol !== 'undefined' && Symbol.toStringTag) {
/******/ 				Object.defineProperty(exports, Symbol.toStringTag, { value: 'Module' });
/******/ 			}
/******/ 			Object.defineProperty(exports, '__esModule', { value: true });
/******/ 		};
/******/ 	})();
/******/ 	
/************************************************************************/
/******/ 	
/******/ 	// startup
/******/ 	// Load entry module and return exports
/******/ 	// This entry module can't be inlined because the eval devtool is used.
/******/ 	var __webpack_exports__ = {};
/******/ 	__webpack_modules__["./InputDateFieldColorChange/index.ts"](0, __webpack_exports__, __webpack_require__);
/******/ 	pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad = __webpack_exports__;
/******/ 	
/******/ })()
;
if (window.ComponentFramework && window.ComponentFramework.registerControl) {
	ComponentFramework.registerControl('InputDateFieldColorChange.InputDateFieldToChange', pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad.InputDateFieldToChange);
} else {
	var InputDateFieldColorChange = InputDateFieldColorChange || {};
	InputDateFieldColorChange.InputDateFieldToChange = pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad.InputDateFieldToChange;
	pcf_tools_652ac3f36e1e4bca82eb3c1dc44e6fad = undefined;
}