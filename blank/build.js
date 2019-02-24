'use strict';

const path = require('path');
const fs = require('fs');

const blankPath = path.join(__dirname, 'blank.xlsx');
const templatePath = path.join(__dirname, 'template.ts');
const buildPath = path.join(__dirname, '..', 'src', 'blank.ts');

const blank = fs.readFileSync(blankPath, 'base64');
const template = fs.readFileSync(templatePath, 'utf8');
const output = template.replace('{{DATA}}', blank);
fs.writeFileSync(buildPath, output);
