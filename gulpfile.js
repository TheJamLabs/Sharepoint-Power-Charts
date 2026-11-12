'use strict';

const build = require('@microsoft/sp-build-web');

build.addSuppression(`Warning - [sass] The local CSS class 'ms-Grid' is not used`);
build.addSuppression(`Warning - [sass] The local CSS class 'ms-fontColor-themePrimary' is not used`);

build.initialize(require('gulp'));