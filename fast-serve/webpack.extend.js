const path = require('path');

module.exports = {
  /**
   * Use this section to transform the existing webpack config.
   * @param config the existing webpack config
   * @param webpack the webpack object
   * @returns the transformed webpack config
   */
  transformConfig: function (config, webpack) {
    config.plugins.push(
      new webpack.NormalModuleReplacementPlugin(
        /ErrorMessage\.module\.scss/,
        path.resolve(__dirname, 'empty-module.js')
      )
    );
    config.plugins.push(
      new webpack.NormalModuleReplacementPlugin(
        /PeoplePickerComponent\.module\.scss/,
        path.resolve(__dirname, 'empty-module.js')
      )
    );
    return config;
  },
  /**
   * Use this section to merge your own webpack config into the existing one.
   */
  webpackConfig: {
  }
};
