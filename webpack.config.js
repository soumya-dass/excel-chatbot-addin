/* eslint-disable no-undef */

// Load .env variables into process.env before Webpack config is evaluated
require('dotenv').config();

const webpack = require('webpack');
const devCerts = require("office-addin-dev-certs");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");

const urlDev = "https://localhost:3000/";
// Ensure this matches your GitHub Pages *base* URL including the repository name and /dist/
// If your GitHub Pages serves directly from the root of the branch (e.g., gh-pages branch containing dist content),
// you might need "https://soumya-dass.github.io/excel-chatbot-addin/"
const urlProd = "https://soumya-dass.github.io/excel-chatbot-addin/dist/"; 

async function getHttpsOptions() {
  const httpsOptions = await devCerts.getHttpsServerOptions();
  return { ca: httpsOptions.ca, key: httpsOptions.key, cert: httpsOptions.cert };
}

module.exports = async (env, options) => {
  const dev = options.mode === "development";
  const config = {
    devtool: "source-map",
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      // Only include the main JS file for taskpane. HtmlWebpackPlugin handles taskpane.html
      taskpane: ["./src/taskpane/taskpane.js"], 
      commands: "./src/commands/commands.js",
    },
    output: {
      clean: true,
      // Specify the output directory for production builds (like GitHub Pages)
      path: require("path").resolve(__dirname, "dist"), 
    },
    resolve: {
      extensions: [".html", ".js"],
    },
    module: {
      rules: [
        // --- CORRECTED & COMBINED JAVASCRIPT/BABEL RULE ---
        {
          test: /\.js$/,
          exclude: /node_modules/, // Don't process files in node_modules
          use: {
            loader: "babel-loader",
            options: {
              presets: [
                ["@babel/preset-env", {
                  // Default behavior of preset-env should correctly transform ES modules
                  // for Webpack's consumption. No special 'modules' option needed here.
                }]
              ],
            },
          },
        },
        // --- HTML LOADER RULE ---
        {
          test: /\.html$/,
          exclude: /node_modules/,
          use: "html-loader",
        },
        // --- IMAGE ASSET RULE ---
        {
          test: /\.(png|jpg|jpeg|gif|ico)$/,
          type: "asset/resource",
          generator: {
            filename: "assets/[name][ext][query]",
          },
        },
        // --- CSS LOADER RULE (THIS WAS MISSING!) ---
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [ // <--- NO EXTRA ARRAY BRACKETS HERE
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]" + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      // --- API Key handling (removed new Dotenv() plugin) ---
      new webpack.DefinePlugin({
        // This will replace 'process.env.GEMINI_API_KEY' in your client-side code
        // with the actual string value from your .env file.
        'process.env.GEMINI_API_KEY': JSON.stringify(process.env.GEMINI_API_KEY || '')
      }),
    ],
    devServer: {
      headers: {
        "Access-Control-Allow-Origin": "*",
      },
      server: {
        type: "https",
        options: env.WEBPACK_BUILD || options.https !== undefined ? options.https : await getHttpsOptions(),
      },
      port: process.env.npm_package_config_dev_server_port || 3000,
    },
  };

  return config;
};