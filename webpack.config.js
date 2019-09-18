const devCerts = require("office-addin-dev-certs");
const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const ExtractTextPlugin = require('extract-text-webpack-plugin');
const HtmlWebpackPlugin = require("html-webpack-plugin");
const webpack = require('webpack');
const MomentLocalesPlugin = require('moment-locales-webpack-plugin');

module.exports = async (env, options) => {
    const dev = options.mode === "development";
    return {
        devtool: dev === "development" ? "source-map" : false,
        entry: {
            // vendor: [
            //     'react',
            //     'react-dom',
            //     'core-js',
            //     'office-ui-fabric-react',
            //     'moment'
            // ],
            //polyfill: 'babel-polyfill',
            command: [
                'babel-polyfill',
                './src/command/command.js'
            ],
            dialog: [
                'babel-polyfill',
                'react-hot-loader/patch',
                './src/dialog/dialog.js'
            ]
        },
        resolve: {
            extensions: [".ts", ".tsx", ".html", ".js"]
        },
        // optimization: {
        //     splitChunks: {
        //         chunks: 'all'
        //     }
        // },
        module: {
            rules: [
                {
                    test: /\.jsx?$/,
                    use: [
                        'react-hot-loader/webpack',
                        'babel-loader',
                    ],
                    exclude: /node_modules/
                },
                {
                    test: /\.css$/,
                    use: ['style-loader', 'css-loader']
                },
                {
                    test: /\.(png|jpe?g|gif|svg|woff|woff2|ttf|eot|ico)$/,
                    use: {
                        loader: 'file-loader',
                        query: {
                            name: 'assets/[name].[ext]'
                        }
                    }
                }
            ]
        },
        plugins: [
            new CleanWebpackPlugin(),
            new CopyWebpackPlugin([
                {
                    to: "dialog.css",
                    from: "./src/dialog/dialog.css"
                }
            ]),
            new ExtractTextPlugin('[name].[hash].css'),
            new HtmlWebpackPlugin({
                filename: "command.html",
                template: './src/command/command.html',
                inject: false,
                //chunks: ['vendor', 'polyfill','command']
                chunks: ['command']
            }),
            new HtmlWebpackPlugin({
                filename: "dialog.html",
                template: "./src/dialog/dialog.html",
                inject: false,
                //chunks: ['vendor', 'polyfill','dialog']
                chunks: ['dialog']
            }),
            new CopyWebpackPlugin([
                {
                    from: './assets',
                    ignore: ['*.scss'],
                    to: 'assets',
                }
            ]),
            new webpack.ProvidePlugin({
                Promise: ["es6-promise", "Promise"]
            }),
            new MomentLocalesPlugin()
        ],
        devServer: {
            headers: {
                "Access-Control-Allow-Origin": "*"
            },
            https: (options.https !== undefined) ? options.https : await devCerts.getHttpsServerOptions(),
            port: process.env.npm_package_config_dev_server_port || 3000
        }
    };
};
