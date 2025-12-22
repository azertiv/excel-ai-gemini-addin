const path = require("path");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const devCerts = require("office-addin-dev-certs");

module.exports = async () => {
  const httpsOptions = await devCerts.getHttpsServerOptions();

  return {
    devtool: "source-map",
    entry: {
      taskpane: "./src/taskpane/taskpane.js",
      functions: "./src/functions/functions.js"
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].js",
      clean: true
    },
    resolve: {
      extensions: [".js"]
    },
    module: {
      rules: [
        {
          test: /\.js$/,
          exclude: /node_modules/,
          use: {
            loader: "babel-loader"
          }
        }
      ]
    },
    plugins: [
      new CopyWebpackPlugin({
        patterns: [
          { from: "src/taskpane/taskpane.html", to: "taskpane.html" },
          { from: "src/taskpane/taskpane.css", to: "taskpane.css" },
          { from: "src/functions/functions.json", to: "functions.json" },
          { from: "assets", to: "assets", noErrorOnMissing: true }
        ]
      })
    ],
    devServer: {
      port: 3000,
      https: httpsOptions,
      headers: { "Access-Control-Allow-Origin": "*" },
      static: { directory: path.join(__dirname, "dist") },
      allowedHosts: "all",
      hot: false,
      client: { overlay: false }
    }
  };
};
