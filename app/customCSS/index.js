const fs = require("node:fs");

exports.onDidFinishLoad = function onDidFinishLoad(content, config) {
  const customCssLocation = getCustomCssLocation(config);
  if (customCssLocation) {
    applyCustomCSSToContent(content, customCssLocation);
  }
};

function getCustomCssLocation(config) {
  return config.customCSSLocation || null;
}

function applyCustomCSSToContent(content, cssLocation) {
  fs.readFile(cssLocation, "utf-8", (error, data) => {
    if (!error) {
      content.insertCSS(data);
    }
  });
}
