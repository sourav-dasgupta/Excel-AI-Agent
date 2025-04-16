const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const webpack = require('webpack');
const dotenv = require('dotenv');

// Load environment variables from .env file
dotenv.config();

module.exports = {
  mode: 'development',
  entry: {
    taskpane: "./src/taskpane/taskpane.ts"
  },
  output: {
    filename: "[name].js"
  },
  resolve: {
    extensions: [".ts", ".tsx", ".html", ".js"]
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        exclude: /node_modules/,
        use: "babel-loader"
      },
      {
        test: /\.html$/,
        use: "html-loader"
      },
      {
        test: /\.css$/,
        use: ["style-loader", "css-loader"]
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: "./src/taskpane/taskpane.html",
      chunks: ["taskpane"]
    }),
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "assets",
          to: "assets"
        },
        {
          from: "manifest*.xml",
          to: "[name][ext]"
        }
      ]
    }),
    new webpack.DefinePlugin({
      // Define process.env object with environment variables
      'process.env': JSON.stringify({
        OPENAI_API_KEY: process.env.OPENAI_API_KEY || ''
      })
    })
  ],
  devServer: {
    headers: { "Access-Control-Allow-Origin": "*" },
    https: true,
    port: 3000
  }
}; 