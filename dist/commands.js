!function(){var e={};e.g=function(){if("object"==typeof globalThis)return globalThis;try{return this||new Function("return this")()}catch(e){if("object"==typeof window)return window}}();var o={};!function(e,t,i){console.log("COMMANDS . TS"),Office.onReady((function(){console.log("OFFICE . ON READY")})),("undefined"!=typeof self?self:"undefined"!=typeof window?window:void 0!==i.g?i.g:void 0).action=function(e){var o={type:Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,message:"Performed action.",icon:"Icon.80x80",persistent:!0};Office.context.mailbox.item.notificationMessages.replaceAsync("action",o),e.completed()},function(){var e="undefined"!=typeof reactHotLoaderGlobal?reactHotLoaderGlobal.default:void 0;if(e){var i=void 0!==o?o:t;if(i)if("function"!=typeof i){for(var n in i)if(Object.prototype.hasOwnProperty.call(i,n)){var c=void 0;try{c=i[n]}catch(e){continue}e.register(c,n,"/Users/admin/code/editor-poc/src/addin/commands.ts")}}else e.register(i,"module.exports","/Users/admin/code/editor-poc/src/addin/commands.ts")}}()}(0,o,e)}();
//# sourceMappingURL=commands.js.map