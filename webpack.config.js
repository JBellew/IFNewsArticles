const path = require('path');
const MiniCssExtractPlugin = require('mini-css-extract-plugin');
const PowerBICustomVisualsWebpackPlugin = require('powerbi-visuals-webpack-plugin');
const packageJson = require('./package.json');
const pbivizJson = require('./pbiviz.json');

module.exports = {
    entry: './src/visual.ts',
    output: {
        path: path.resolve(__dirname, 'dist'),
        filename: 'visual.js',
        library: 'powerbiVisuals',
        libraryTarget: 'var',
    },
    module: {
        rules: [
            {
                test: /\.ts(x?)$/,
                loader: 'ts-loader',
                exclude: /node_modules/,
            },
            {
                test: /\.css$/,
                use: [MiniCssExtractPlugin.loader, 'css-loader'],
            },
            {
                test: /\.(woff|woff2|eot|ttf|svg)$/,
                loader: 'file-loader',
                options: {
                    name: '[name].[ext]',
                    outputPath: 'fonts/'
                }
            },
        ],
    },
    resolve: {
        extensions: ['.ts', '.tsx', '.js'],
    },
    plugins: [
        new MiniCssExtractPlugin({
            filename: 'visual.css',
        }),
        new PowerBICustomVisualsWebpackPlugin({
            pbivizJson: pbivizJson,
            capabilitiesJson: pbivizJson.capabilities,
            stringResources: [],
            jsonResources: [],
            pluginName: packageJson.name,
        }),
    ],
    externals: {
        'powerbi-visuals-api': 'null',
    },
};
