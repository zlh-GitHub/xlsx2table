import path from 'path';
import commonjs from '@rollup/plugin-commonjs';
import json from '@rollup/plugin-json';
import shebang from 'rollup-plugin-preserve-shebang';
import babel from '@rollup/plugin-babel';
import copy from 'rollup-plugin-copy'

export default [
  {
    input: path.resolve(__dirname, './index.js'),
    output: {
      dir: './lib',
      format: 'cjs',
    },
    plugins: [
      shebang({ shebang: '#!/usr/bin/env node' }),
      commonjs(),
      babel({
        exclude: 'node_modules/**',
        // babelHelpers: 'runtime',
      }),
      json(),
      copy({
        targets: [
          { src: ['./public', './config.json'], dest: 'lib/' },
        ]
      })
    ],
  },
];
