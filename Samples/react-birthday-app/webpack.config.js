'use strict';

/** Note: this require may need to be fixed to point to the build that exports the gulp-core-build-webpack instance. */
let webpackTaskResources = require('@microsoft/web-library-build').webpack.resources;
let webpack = webpackTaskResources.webpack;

let path = require('path');
let isProduction = process.argv.indexOf('--production') > -1;
let packageJSON = require('./package.json');

let webpackConfig = {
    context: path.join(__dirname, 'lib/'),

    entry: {
        [packageJSON.name]: './index.js'
    },

    output: {
        libraryTarget: 'umd',
        path: path.join(__dirname, '/dist'),
        filename: `[name]${ isProduction ? '.min' : '' }.js`
    },

    devtool: 'source-map',

    devServer: {
        stats: 'none'
    },

    module: {
        loaders: [
            { test: /\.css$/, loader: "style!css!" }
        ]
    },

    externals: {
        "react": "React",
        "react-dom": "ReactDOM",
        "sp-init": {
            "path": "http://sp13dev:81/sites/zerhusen/_layouts/15/init.js",
            "globalName": "$_global_init"
        },
        "microsoft-ajax": {
            "path": "http://sp13dev:81/sites/zerhusen/_layouts/15/MicrosoftAjax.js",
            "globalName": "Sys",
            "globalDependencies": [
                "sp-init"
            ]
        },
        "sp-runtime": {
            "path": "http://sp13dev:81/sites/zerhusen/_layouts/15/SP.Runtime.js",
            "globalName": "SP",
            "globalDependencies": [
                "microsoft-ajax"
            ]
        },
        "sharepoint": {
            "path": "http://sp13dev:81/sites/zerhusen/_layouts/15/SP.js",
            "globalName": "SP",
            "globalDependencies": [
                "sp-runtime"
            ]
        }
    },

    plugins: [
        //  new WebpackNotifierPlugin()
    ]
};

if (isProduction) {
    webpackConfig.plugins.push(new webpack.optimize.UglifyJsPlugin({
        minimize: true,
        compress: {
            warnings: false
        }
    }));
}

module.exports = webpackConfig;