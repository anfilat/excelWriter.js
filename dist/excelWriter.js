/*
excelWriter - Javascript module for generating Excel files
<https://github.com/anfilat/excelWriter.js>
(c) 2017 Andrey Filatkin
Dual licenced under the MIT license or GPLv3
*/
(function(f){if(typeof exports==="object"&&typeof module!=="undefined"){module.exports=f()}else if(typeof define==="function"&&define.amd){define([],f)}else{var g;if(typeof window!=="undefined"){g=window}else if(typeof global!=="undefined"){g=global}else if(typeof self!=="undefined"){g=self}else{g=this}g.excelWriter = f()}})(function(){var define,module,exports;return (function e(t,n,r){function s(o,u){if(!n[o]){if(!t[o]){var a=typeof require=="function"&&require;if(!u&&a)return a(o,!0);if(i)return i(o,!0);var f=new Error("Cannot find module '"+o+"'");throw f.code="MODULE_NOT_FOUND",f}var l=n[o]={exports:{}};t[o][0].call(l.exports,function(e){var n=t[o][1][e];return s(n?n:e)},l,l.exports,e,t,n,r)}return n[o].exports}var i=typeof require=="function"&&require;for(var o=0;o<r.length;o++)s(r[o]);return s})({1:[function(require,module,exports){

},{}],2:[function(require,module,exports){
(function (global){
'use strict';

var _slicedToArray = function () { function sliceIterator(arr, i) { var _arr = []; var _n = true; var _d = false; var _e = undefined; try { for (var _i = arr[Symbol.iterator](), _s; !(_n = (_s = _i.next()).done); _n = true) { _arr.push(_s.value); if (i && _arr.length === i) break; } } catch (err) { _d = true; _e = err; } finally { try { if (!_n && _i["return"]) _i["return"](); } finally { if (_d) throw _e; } } return _arr; } return function (arr, i) { if (Array.isArray(arr)) { return arr; } else if (Symbol.iterator in Object(arr)) { return sliceIterator(arr, i); } else { throw new TypeError("Invalid attempt to destructure non-iterable instance"); } }; }();

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');

/**
 * @param {{
 *  name: string,
 *  value?: *,
 *  children?: Array.<string>,
 *  attributes?: Array.<string, *>,
 *  ns?: string
 * }} config
 */
function toXMLString(config) {
	var string = '<' + config.name;
	var content = '';

	if (config.ns) {
		string = util.xmlPrefix + string + ' xmlns="' + util.schemas[config.ns] + '"';
	}
	if (config.attributes) {
		config.attributes.forEach(function (_ref) {
			var _ref2 = _slicedToArray(_ref, 2),
			    key = _ref2[0],
			    value = _ref2[1];

			string += ' ' + key + '="' + _.escape(value) + '"';
		});
	}
	if (config.value !== undefined) {
		content += _.escape(config.value);
	}
	if (config.children) {
		content += config.children.join('');
	}
	if (content) {
		string += '>' + content + '</' + config.name + '>';
	} else {
		string += '/>';
	}

	return string;
}

module.exports = toXMLString;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./util":26}],3:[function(require,module,exports){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Images = function () {
	function Images(common) {
		_classCallCheck(this, Images);

		this.common = common;
		this._images = Object.create(null);
		this._imageByNames = Object.create(null);
		this._extensions = Object.create(null);
	}

	_createClass(Images, [{
		key: 'addImage',
		value: function addImage(data) {
			var type = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : '';
			var name = arguments[2];

			var image = this._images[data];

			if (!image) {
				var id = this.common.uniqueIdForSpace('image');
				var path = 'xl/media/image' + id + '.' + type;
				var contentType = void 0;

				name = name || 'excelWriter' + id;
				switch (type.toLowerCase()) {
					case 'jpeg':
					case 'jpg':
						contentType = 'image/jpeg';
						break;
					case 'png':
						contentType = 'image/png';
						break;
					case 'gif':
						contentType = 'image/gif';
						break;
					default:
						contentType = null;
						break;
				}

				image = {
					objectId: 'image' + id,
					data: data,
					name: name,
					contentType: contentType,
					extension: type,
					path: path
				};
				this.common.addPath(image, '/' + path);
				this._images[data] = image;
				this._imageByNames[name] = image;
				this._extensions[type] = contentType;
			} else if (name && !this._imageByNames[name]) {
				image.name = name;
				this._imageByNames[name] = image;
			}
			return image.name;
		}
	}, {
		key: 'getImage',
		value: function getImage(name) {
			return this._imageByNames[name];
		}
	}, {
		key: 'getImages',
		value: function getImages() {
			return this._images;
		}
	}, {
		key: 'removeImages',
		value: function removeImages() {
			this._images = null;
			this._imageByNames = null;
		}
	}, {
		key: 'getExtensions',
		value: function getExtensions() {
			return this._extensions;
		}
	}]);

	return Images;
}();

module.exports = Images;

},{}],4:[function(require,module,exports){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Images = require('./images');
var SharedStrings = require('./sharedStrings');
var Styles = require('../styles');

var Common = function () {
	function Common() {
		_classCallCheck(this, Common);

		this.idSpaces = Object.create(null);
		this._paths = Object.create(null);

		this.images = new Images(this);

		this.strings = new SharedStrings(this);
		this.addPath(this.strings, 'sharedStrings.xml');

		this.styles = new Styles(this);
		this.addPath(this.styles, 'styles.xml');

		this.worksheets = [];
		this.tables = [];
		this.drawings = [];
	}

	_createClass(Common, [{
		key: 'uniqueId',
		value: function uniqueId(space) {
			return space + this.uniqueIdForSpace(space);
		}
	}, {
		key: 'uniqueIdForSpace',
		value: function uniqueIdForSpace(space) {
			if (!this.idSpaces[space]) {
				this.idSpaces[space] = 1;
			}
			return this.idSpaces[space]++;
		}
	}, {
		key: 'addPath',
		value: function addPath(object, path) {
			this._paths[object.objectId] = path;
		}
	}, {
		key: 'getPath',
		value: function getPath(object) {
			return this._paths[object.objectId];
		}
	}, {
		key: 'addWorksheet',
		value: function addWorksheet(worksheet) {
			var index = this.worksheets.length + 1;
			var path = 'worksheets/sheet' + index + '.xml';
			var relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

			worksheet.path = 'xl/' + path;
			worksheet.relationsPath = relationsPath;
			this.worksheets.push(worksheet);
			this.addPath(worksheet, path);
		}
	}, {
		key: 'getNewWorksheetDefaultName',
		value: function getNewWorksheetDefaultName() {
			return 'Sheet ' + (this.worksheets.length + 1);
		}
	}, {
		key: 'setActiveWorksheet',
		value: function setActiveWorksheet(worksheet) {
			this.activeWorksheet = worksheet;
		}
	}, {
		key: 'addTable',
		value: function addTable(table) {
			var index = this.tables.length + 1;
			var path = 'xl/tables/table' + index + '.xml';

			table.path = path;
			this.tables.push(table);
			this.addPath(table, '/' + path);
		}
	}, {
		key: 'addDrawings',
		value: function addDrawings(drawings) {
			var index = this.drawings.length + 1;
			var path = 'xl/drawings/drawing' + index + '.xml';
			var relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

			drawings.path = path;
			drawings.relationsPath = relationsPath;
			this.drawings.push(drawings);
			this.addPath(drawings, '/' + path);
		}
	}]);

	return Common;
}();

module.exports = Common;

},{"../styles":18,"./images":3,"./sharedStrings":5}],5:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Readable = require('stream').Readable;
var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');

var spaceRE = /^\s|\s$/;

var SharedStrings = function () {
	function SharedStrings(common) {
		_classCallCheck(this, SharedStrings);

		this.objectId = common.uniqueId('SharedStrings');
		this._strings = Object.create(null);
		this._stringArray = [];
		this.count = 0;
	}

	_createClass(SharedStrings, [{
		key: 'add',
		value: function add(string) {
			var stringId = this._strings[string];

			if (stringId === undefined) {
				stringId = this.count++;

				this._strings[string] = stringId;
				this._stringArray[stringId] = string;
			}

			return stringId;
		}
	}, {
		key: 'isStrings',
		value: function isStrings() {
			return this.count > 0;
		}
	}, {
		key: 'save',
		value: function save(canStream) {
			this._strings = null;

			if (canStream) {
				return new SharedStringsStream({
					strings: this._stringArray
				});
			}
			var result = getXMLBegin(this.count) + this._stringArray.map(function (string) {
				string = _.escape(string);

				if (spaceRE.test(string)) {
					return '<si><t xml:space="preserve">' + string + '</t></si>';
				}
				return '<si><t>' + string + '</t></si>';
			}).join('') + getXMLEnd();
			this._stringArray = null;
			return result;
		}
	}]);

	return SharedStrings;
}();

var SharedStringsStream = function (_ref) {
	_inherits(SharedStringsStream, _ref);

	function SharedStringsStream(options) {
		_classCallCheck(this, SharedStringsStream);

		var _this = _possibleConstructorReturn(this, (SharedStringsStream.__proto__ || Object.getPrototypeOf(SharedStringsStream)).call(this, options));

		_this.strings = options.strings;
		_this.status = 0;
		return _this;
	}

	_createClass(SharedStringsStream, [{
		key: '_read',
		value: function _read(size) {
			var stop = false;

			if (this.status === 0) {
				stop = !this.push(getXMLBegin(this.strings.length));

				this.status = 1;
				this.index = 0;
				this.len = this.strings.length;
			}

			if (this.status === 1) {
				var s = '';
				var str = void 0;

				while (this.index < this.len && !stop) {
					while (this.index < this.len && s.length < size) {
						str = _.escape(this.strings[this.index]);

						if (spaceRE.test(str)) {
							s += '<si><t xml:space="preserve">' + str + '</t></si>';
						} else {
							s += '<si><t>' + str + '</t></si>';
						}
						this.strings[this.index] = null;
						this.index++;
					}
					stop = !this.push(s);
					s = '';
				}

				if (this.index === this.len) {
					this.status = 2;
				}
			}

			if (this.status === 2) {
				this.push(getXMLEnd());
				this.push(null);
			}
		}
	}]);

	return SharedStringsStream;
}(Readable || null);

function getXMLBegin(length) {
	return util.xmlPrefix + '<sst xmlns="' + util.schemas.spreadsheetml + '" count="' + length + '" uniqueCount="' + length + '">';
}

function getXMLEnd() {
	return '</sst>';
}

module.exports = SharedStrings;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../util":26,"stream":1}],6:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

var Anchor = function () {
	function Anchor(config) {
		_classCallCheck(this, Anchor);

		var coord = void 0;

		if (_.isObject(config)) {
			if (_.has(config, 'cell')) {
				if (_.isObject(config.cell)) {
					coord = { x: config.cell.c || 1, y: config.cell.r || 1 };
				} else {
					coord = util.letterToPosition(config.cell || '');
				}
			} else {
				coord = { x: config.c || 1, y: config.r || 1 };
			}
		} else {
			coord = util.letterToPosition(config || '');
			config = {};
		}
		var x = coord.x - 1;
		var y = coord.y - 1;

		this.from = {
			x: x,
			y: y,
			xOff: util.pixelsToEMUs(config.left || 0),
			yOff: util.pixelsToEMUs(config.top || 0)
		};
		this.to = {
			x: x + (config.cols || 1),
			y: y + (config.rows || 1),
			xOff: util.pixelsToEMUs(-config.right || 0),
			yOff: util.pixelsToEMUs(-config.bottom || 0)
		};
	}

	_createClass(Anchor, [{
		key: 'saveWithContent',
		value: function saveWithContent(content) {
			return toXMLString({
				name: 'xdr:twoCellAnchor',
				children: [toXMLString({
					name: 'xdr:from',
					children: [toXMLString({
						name: 'xdr:col',
						value: this.from.x
					}), toXMLString({
						name: 'xdr:colOff',
						value: this.from.xOff
					}), toXMLString({
						name: 'xdr:row',
						value: this.from.y
					}), toXMLString({
						name: 'xdr:rowOff',
						value: this.from.yOff
					})]
				}), toXMLString({
					name: 'xdr:to',
					children: [toXMLString({
						name: 'xdr:col',
						value: this.to.x
					}), toXMLString({
						name: 'xdr:colOff',
						value: this.to.xOff
					}), toXMLString({
						name: 'xdr:row',
						value: this.to.y
					}), toXMLString({
						name: 'xdr:rowOff',
						value: this.to.yOff
					})]
				}), content, toXMLString({
					name: 'xdr:clientData'
				})]
			});
		}
	}]);

	return Anchor;
}();

module.exports = Anchor;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":26}],7:[function(require,module,exports){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var util = require('../util');
var toXMLString = require('../XMLString');

var AnchorAbsolute = function () {
	function AnchorAbsolute() {
		var _ref = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {},
		    _ref$left = _ref.left,
		    left = _ref$left === undefined ? 0 : _ref$left,
		    _ref$top = _ref.top,
		    top = _ref$top === undefined ? 0 : _ref$top,
		    _ref$width = _ref.width,
		    width = _ref$width === undefined ? 0 : _ref$width,
		    _ref$height = _ref.height,
		    height = _ref$height === undefined ? 0 : _ref$height;

		_classCallCheck(this, AnchorAbsolute);

		this.x = util.pixelsToEMUs(left);
		this.y = util.pixelsToEMUs(top);
		this.width = util.pixelsToEMUs(width);
		this.height = util.pixelsToEMUs(height);
	}

	_createClass(AnchorAbsolute, [{
		key: 'saveWithContent',
		value: function saveWithContent(content) {
			return toXMLString({
				name: 'xdr:absoluteAnchor',
				children: [toXMLString({
					name: 'xdr:pos',
					attributes: [['x', this.x], ['y', this.y]]
				}), toXMLString({
					name: 'xdr:ext',
					attributes: [['cx', this.width], ['cy', this.height]]
				}), content, toXMLString({
					name: 'xdr:clientData'
				})]
			});
		}
	}]);

	return AnchorAbsolute;
}();

module.exports = AnchorAbsolute;

},{"../XMLString":2,"../util":26}],8:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

var AnchorOneCell = function () {
	function AnchorOneCell(config) {
		_classCallCheck(this, AnchorOneCell);

		var coord = void 0;

		if (_.isObject(config)) {
			if (_.has(config, 'cell')) {
				if (_.isObject(config.cell)) {
					coord = { x: config.cell.c || 1, y: config.cell.r || 1 };
				} else {
					coord = util.letterToPosition(config.cell || '');
				}
			} else {
				coord = { x: config.c || 1, y: config.r || 1 };
			}
		} else {
			coord = util.letterToPosition(config || '');
			config = {};
		}

		this.x = coord.x - 1;
		this.y = coord.y - 1;
		this.xOff = util.pixelsToEMUs(config.left || 0);
		this.yOff = util.pixelsToEMUs(config.top || 0);
		this.width = util.pixelsToEMUs(config.width || 0);
		this.height = util.pixelsToEMUs(config.height || 0);
	}

	_createClass(AnchorOneCell, [{
		key: 'saveWithContent',
		value: function saveWithContent(content) {
			return toXMLString({
				name: 'xdr:oneCellAnchor',
				children: [toXMLString({
					name: 'xdr:from',
					children: [toXMLString({
						name: 'xdr:col',
						value: this.x
					}), toXMLString({
						name: 'xdr:colOff',
						value: this.xOff
					}), toXMLString({
						name: 'xdr:row',
						value: this.y
					}), toXMLString({
						name: 'xdr:rowOff',
						value: this.yOff
					})]
				}), toXMLString({
					name: 'xdr:ext',
					attributes: [['cx', this.width], ['cy', this.height]]
				}), content, toXMLString({
					name: 'xdr:clientData'
				})]
			});
		}
	}]);

	return AnchorOneCell;
}();

module.exports = AnchorOneCell;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":26}],9:[function(require,module,exports){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var Relations = require('../relations');
var util = require('../util');
var toXMLString = require('../XMLString');
var Picture = require('./picture');

var Drawings = function () {
	function Drawings(common) {
		_classCallCheck(this, Drawings);

		this.common = common;

		this.objectId = this.common.uniqueId('Drawings');
		this.drawings = [];
		this.relations = new Relations(common);
	}

	_createClass(Drawings, [{
		key: 'addImage',
		value: function addImage(name, config, anchorType) {
			var image = this.common.images.getImage(name);
			var imageRelationId = this.relations.addRelation(image, 'image');
			var picture = new Picture(this.common, {
				image: image,
				imageRelationId: imageRelationId,
				config: config,
				anchorType: anchorType
			});

			this.drawings.push(picture);
		}
	}, {
		key: 'save',
		value: function save() {
			var attributes = [['xmlns:a', util.schemas.drawing], ['xmlns:r', util.schemas.relationships], ['xmlns:xdr', util.schemas.spreadsheetDrawing]];
			var children = this.drawings.map(function (picture) {
				return picture.save();
			});

			return toXMLString({
				name: 'xdr:wsDr',
				ns: 'spreadsheetDrawing',
				attributes: attributes,
				children: children
			});
		}
	}]);

	return Drawings;
}();

module.exports = Drawings;

},{"../XMLString":2,"../relations":12,"../util":26,"./picture":10}],10:[function(require,module,exports){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var util = require('../util');
var toXMLString = require('../XMLString');
var Anchor = require('./anchor');
var AnchorOneCell = require('./anchorOneCell');
var AnchorAbsolute = require('./anchorAbsolute');

var Picture = function () {
	function Picture(common, config) {
		_classCallCheck(this, Picture);

		this.pictureId = common.uniqueIdForSpace('Picture');
		this.image = config.image;
		this.imageRelationId = config.imageRelationId;
		this.createAnchor(config.anchorType, config.config);
	}
	/**
  *
  * @param {String} type Can be 'anchor', 'oneCell' or 'absolute'.
  * @param {Object} config Shorthand - pass the created anchor coords that can normally be used to construct it.
  */


	_createClass(Picture, [{
		key: 'createAnchor',
		value: function createAnchor(type, config) {
			switch (type) {
				case 'anchor':
					this.anchor = new Anchor(config);
					break;
				case 'oneCell':
					this.anchor = new AnchorOneCell(config);
					break;
				case 'absolute':
					this.anchor = new AnchorAbsolute(config);
					break;
			}
		}
	}, {
		key: 'save',
		value: function save() {
			return this.anchor.saveWithContent(toXMLString({
				name: 'xdr:pic',
				children: [toXMLString({
					name: 'xdr:nvPicPr',
					children: [toXMLString({
						name: 'xdr:cNvPr',
						attributes: [['id', this.pictureId], ['name', this.image.name]]
					}), toXMLString({
						name: 'xdr:cNvPicPr',
						children: [toXMLString({
							name: 'a:picLocks',
							attributes: [['noChangeAspect', '1'], ['noChangeArrowheads', '1']]
						})]
					})]
				}), toXMLString({
					name: 'xdr:blipFill',
					children: [toXMLString({
						name: 'a:blip',
						attributes: [['xmlns:r', util.schemas.relationships], ['r:embed', this.imageRelationId]]
					}), toXMLString({
						name: 'a:srcRect'
					}), toXMLString({
						name: 'a:stretch',
						children: [toXMLString({
							name: 'a:fillRect'
						})]
					})]
				}), toXMLString({
					name: 'xdr:spPr',
					attributes: [['bwMode', 'auto']],
					children: [toXMLString({
						name: 'a:xfrm'
					}), toXMLString({
						name: 'a:prstGeom',
						attributes: [['prst', 'rect']]
					})]
				})]
			}));
		}
	}]);

	return Picture;
}();

module.exports = Picture;

},{"../XMLString":2,"../util":26,"./anchor":6,"./anchorAbsolute":7,"./anchorOneCell":8}],11:[function(require,module,exports){
(function (global){
'use strict';

var Readable = require('stream').Readable;
var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var JSZip = typeof window !== "undefined" ? window['JSZip'] : typeof global !== "undefined" ? global['JSZip'] : null;
var Workbook = require('./workbook');

module.exports = {
	createWorkbook: function createWorkbook() {
		return new Workbook();
	},

	/**
  * Turns a workbook into a downloadable file.
  * @param {Workbook} workbook - The workbook that is being converted
  * @param {Object?} options - options to modify how the zip is created. See http://stuk.github.io/jszip/#doc_generate_options
  */
	save: function save(workbook, options) {
		var zip = new JSZip();
		var canStream = !!Readable;

		workbook._generateFiles(zip, canStream);
		return zip.generateAsync(_.defaults(options, {
			compression: 'DEFLATE',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			type: 'base64'
		}));
	},

	saveAsNodeStream: function saveAsNodeStream(workbook, options) {
		var zip = new JSZip();
		var canStream = !!Readable;

		workbook._generateFiles(zip, canStream);
		return zip.generateNodeStream(_.defaults(options, {
			compression: 'DEFLATE'
		}));
	}
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./workbook":27,"stream":1}],12:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');
var toXMLString = require('./XMLString');

var RelationshipManager = function () {
	function RelationshipManager(common) {
		_classCallCheck(this, RelationshipManager);

		this.common = common;

		this.relations = Object.create(null);
		this.lastId = 1;
	}

	_createClass(RelationshipManager, [{
		key: 'addRelation',
		value: function addRelation(object, type) {
			var relation = this.relations[object.objectId];

			if (relation) {
				return relation.relationId;
			}

			var relationId = 'rId' + this.lastId++;
			this.relations[object.objectId] = {
				relationId: relationId,
				schema: util.schemas[type],
				object: object
			};
			return relationId;
		}
	}, {
		key: 'getRelationshipId',
		value: function getRelationshipId(object) {
			var relation = this.relations[object.objectId];

			return relation ? relation.relationId : null;
		}
	}, {
		key: 'save',
		value: function save() {
			var common = this.common;
			var children = _.map(this.relations, function (relation) {
				var attributes = [['Id', relation.relationId], ['Type', relation.schema], ['Target', relation.object.target || common.getPath(relation.object)]];

				if (relation.object.targetMode) {
					attributes.push(['TargetMode', relation.object.targetMode]);
				}

				return toXMLString({
					name: 'Relationship',
					attributes: attributes
				});
			});

			return toXMLString({
				name: 'Relationships',
				ns: 'relationshipPackage',
				children: children
			});
		}
	}]);

	return RelationshipManager;
}();

module.exports = RelationshipManager;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./util":26}],13:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.alignment.aspx
//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.horizontalalignmentvalues.aspx
var HORIZONTAL = ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.verticalalignmentvalues.aspx
var VERTICAL = ['top', 'center', 'bottom', 'justify', 'distributed'];

function canon(format) {
	var result = {};

	if (_.has(format, 'horizontal') && _.includes(HORIZONTAL, format.horizontal)) {
		result.horizontal = format.horizontal;
	}
	if (_.has(format, 'vertical') && _.includes(VERTICAL, format.vertical)) {
		result.vertical = format.vertical;
	}
	if (_.has(format, 'indent')) {
		result.indent = format.indent;
	}
	if (_.has(format, 'justifyLastLine')) {
		result.justifyLastLine = format.justifyLastLine ? 1 : 0;
	}
	if (_.has(format, 'readingOrder') && _.includes([0, 1, 2], format.readingOrder)) {
		result.readingOrder = format.readingOrder;
	}
	if (_.has(format, 'relativeIndent')) {
		result.relativeIndent = format.relativeIndent;
	}
	if (_.has(format, 'shrinkToFit')) {
		result.shrinkToFit = format.shrinkToFit ? 1 : 0;
	}
	if (_.has(format, 'textRotation')) {
		result.textRotation = format.textRotation;
	}
	if (_.has(format, 'wrapText')) {
		result.wrapText = format.wrapText ? 1 : 0;
	}

	return _.isEmpty(result) ? null : result;
}

function merge(formatTo, formatFrom) {
	return _.assign(formatTo, formatFrom);
}

function saveFormat(format) {
	return toXMLString({
		name: 'alignment',
		attributes: _.toPairs(format)
	});
}

module.exports = {
	canon: canon,
	merge: merge,
	saveFormat: saveFormat
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],14:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');

var _require = require('./utils'),
    saveColor = _require.saveColor;

var toXMLString = require('../XMLString');

var MAIN_BORDERS = ['left', 'right', 'top', 'bottom'];
var BORDERS = ['left', 'right', 'top', 'bottom', 'diagonal'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borders.aspx

var Borders = function (_StylePart) {
	_inherits(Borders, _StylePart);

	function Borders(styles) {
		_classCallCheck(this, Borders);

		var _this = _possibleConstructorReturn(this, (Borders.__proto__ || Object.getPrototypeOf(Borders)).call(this, styles, 'borders', 'border'));

		_this.init();
		_this.lastId = _this.formats.length;
		return _this;
	}

	_createClass(Borders, [{
		key: 'init',
		value: function init() {
			this.formats.push({ format: this.canon({}) });
		}
	}, {
		key: 'canon',
		value: function canon(format) {
			return Borders.canon(format);
		}
	}, {
		key: 'merge',
		value: function merge() {
			var formatTo = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : Borders.canon({});
			var formatFrom = arguments[1];

			if (formatFrom) {
				BORDERS.forEach(function (name) {
					var borderFrom = formatFrom[name];

					if (borderFrom && borderFrom.style) {
						formatTo[name].style = borderFrom.style;
					}
					if (borderFrom && borderFrom.color) {
						formatTo[name].color = borderFrom.color;
					}
				});
			}
			return formatTo;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			return Borders.saveFormat(format, styleFormat);
		}
	}], [{
		key: 'canon',
		value: function canon(format) {
			var result = {};

			if (_.has(format, 'style') || _.has(format, 'color')) {
				MAIN_BORDERS.forEach(function (name) {
					result[name] = {
						style: format.style,
						color: format.color
					};
				});
			} else {
				BORDERS.forEach(function (name) {
					var border = format[name];

					if (border) {
						result[name] = {
							style: border.style,
							color: border.color
						};
					} else {
						result[name] = {};
					}
				});
			}
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var children = BORDERS.map(function (name) {
				var border = format[name];
				var attributes = void 0;
				var children = void 0;

				if (border) {
					if (border.style) {
						attributes = [['style', border.style]];
					}
					if (border.color) {
						children = [saveColor(border.color)];
					}
				}
				return toXMLString({
					name: name,
					attributes: attributes,
					children: children
				});
			});

			return toXMLString({
				name: 'border',
				children: children
			});
		}
	}]);

	return Borders;
}(StylePart);

module.exports = Borders;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./stylePart":21,"./utils":24}],15:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var alignment = require('./alignment');
var protection = require('./protection');
var toXMLString = require('../XMLString');

var ALLOWED_PARTS = ['format', 'fill', 'border', 'font'];
var XLS_NAMES = ['numFmtId', 'fillId', 'borderId', 'fontId'];

var Cells = function (_StylePart) {
	_inherits(Cells, _StylePart);

	function Cells(styles) {
		_classCallCheck(this, Cells);

		var _this = _possibleConstructorReturn(this, (Cells.__proto__ || Object.getPrototypeOf(Cells)).call(this, styles, 'cellXfs', 'format'));

		_this.init();
		_this.lastId = _this.formats.length;
		_this.saveEmpty = false;
		return _this;
	}

	_createClass(Cells, [{
		key: 'init',
		value: function init() {
			this.formats.push({ format: this.canon({}) });
		}
	}, {
		key: 'canon',
		value: function canon(format, flags) {
			var result = {};

			if (format.format) {
				result.format = this.styles.numberFormats.add(format.format);
			}
			if (format.font) {
				result.font = this.styles.fonts.add(format.font);
			}
			if (format.pattern) {
				result.fill = this.styles.fills.add(format.pattern, null, { fillType: 'pattern' });
			} else if (format.gradient) {
				result.fill = this.styles.fills.add(format.gradient, null, { fillType: 'gradient' });
			} else if (flags && flags.merge && format.fill) {
				result.fill = this.styles.fills.add(format.fill, null, flags);
			}
			if (format.border) {
				result.border = this.styles.borders.add(format.border);
			}
			var alignmentValue = alignment.canon(format);
			if (alignmentValue) {
				result.alignment = alignmentValue;
			}
			var protectionValue = protection.canon(format);
			if (protectionValue) {
				result.protection = protectionValue;
			}
			if (format.fillOut) {
				result.fillOut = format.fillOut;
			}
			return result;
		}
	}, {
		key: 'fullGet',
		value: function fullGet(format) {
			if (this.getId(format)) {
				format = this.get(format);
			} else {
				format = this.canon(format);
			}

			var result = {};
			if (format.format) {
				result.format = this.styles.numberFormats.get(format.format);
			}
			if (format.font) {
				result.font = _.clone(this.styles.fonts.get(format.font));
			}
			if (format.fill) {
				result.fill = _.clone(this.styles.fills.get(format.fill));
			}
			if (format.border) {
				result.border = _.clone(this.styles.borders.get(format.border));
			}
			if (format.alignment) {
				result.alignment = _.clone(format.alignment);
			}
			if (format.protection) {
				result.protection = _.clone(format.protection);
			}
			return result;
		}
	}, {
		key: 'cutVisible',
		value: function cutVisible(format) {
			var result = {};

			if (format.format) {
				result.format = format.format;
			}
			if (format.font) {
				result.font = format.font;
			}
			if (format.alignment) {
				result.alignment = format.alignment;
			}
			if (format.protection) {
				result.protection = format.protection;
			}
			return result;
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			if (formatTo.format || formatFrom.format) {
				formatTo.format = this.styles.numberFormats.merge(formatTo.format, formatFrom.format);
			}
			if (formatTo.font || formatFrom.font) {
				formatTo.font = this.styles.fonts.merge(formatTo.font, formatFrom.font);
			}
			if (formatTo.fill || formatFrom.fill) {
				formatTo.fill = this.styles.fills.merge(formatTo.fill, formatFrom.fill);
			}
			if (formatTo.border || formatFrom.border) {
				formatTo.border = this.styles.borders.merge(formatTo.border, formatFrom.border);
			}
			if (formatTo.alignment || formatFrom.alignment) {
				formatTo.alignment = alignment.merge(formatTo.alignment, formatFrom.alignment);
			}
			if (formatTo.protection || formatFrom.protection) {
				formatTo.protection = protection.merge(formatTo.protection, formatFrom.protection);
			}
			return formatTo;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var styles = this.styles;
			var attributes = [];
			var children = [];

			if (format.alignment) {
				children.push(alignment.saveFormat(format.alignment));
				attributes.push(['applyAlignment', 'true']);
			}
			if (format.protection) {
				children.push(protection.saveFormat(format.protection));
				attributes.push(['applyProtection', 'true']);
			}

			_.forEach(format, function (value, key) {
				if (_.includes(ALLOWED_PARTS, key)) {
					var xlsName = XLS_NAMES[_.indexOf(ALLOWED_PARTS, key)];

					if (key === 'format') {
						attributes.push([xlsName, styles.numberFormats.getId(value)]);
						attributes.push(['applyNumberFormat', 'true']);
					} else if (key === 'fill') {
						attributes.push([xlsName, styles.fills.getId(value)]);
						attributes.push(['applyFill', 'true']);
					} else if (key === 'border') {
						attributes.push([xlsName, styles.borders.getId(value)]);
						attributes.push(['applyBorder', 'true']);
					} else if (key === 'font') {
						attributes.push([xlsName, styles.fonts.getId(value)]);
						attributes.push(['applyFont', 'true']);
					}
				}
			});

			return toXMLString({
				name: 'xf',
				attributes: attributes,
				children: children
			});
		}
	}]);

	return Cells;
}(StylePart);

module.exports = Cells;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./alignment":13,"./protection":20,"./stylePart":21}],16:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');

var _require = require('./utils'),
    canonColor = _require.canonColor,
    saveColor = _require.saveColor;

var toXMLString = require('../XMLString');

var PATTERN_TYPES = ['none', 'solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fills.aspx

var Fills = function (_StylePart) {
	_inherits(Fills, _StylePart);

	function Fills(styles) {
		_classCallCheck(this, Fills);

		var _this = _possibleConstructorReturn(this, (Fills.__proto__ || Object.getPrototypeOf(Fills)).call(this, styles, 'fills', 'fill'));

		_this.init();
		_this.lastId = _this.formats.length;
		return _this;
	}

	_createClass(Fills, [{
		key: 'init',
		value: function init() {
			this.formats.push({ format: this.canon({ type: 'none' }, { fillType: 'pattern' }) }, { format: this.canon({ type: 'gray125' }, { fillType: 'pattern' }) });
		}
	}, {
		key: 'canon',
		value: function canon(format, flags) {
			return Fills.canon(format, flags);
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			return formatFrom || formatTo;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			return Fills.saveFormat(format);
		}
	}], [{
		key: 'canon',
		value: function canon(format, flags) {
			var result = {
				fillType: flags.merge ? format.fillType : flags.fillType
			};

			if (result.fillType === 'pattern') {
				var fgColor = (flags.merge ? format.fgColor : format.color) || 'FFFFFFFF';
				var bgColor = (flags.merge ? format.bgColor : format.backColor) || 'FFFFFFFF';
				var patternType = flags.merge ? format.patternType : format.type;

				result.patternType = _.includes(PATTERN_TYPES, patternType) ? patternType : 'solid';
				if (flags.isTable && result.patternType === 'solid') {
					result.fgColor = bgColor;
					result.bgColor = fgColor;
				} else {
					result.fgColor = fgColor;
					result.bgColor = bgColor;
				}
			} else {
				if (_.has(format, 'left')) {
					result.left = format.left || 0;
					result.right = format.right || 0;
					result.top = format.top || 0;
					result.bottom = format.bottom || 0;
				} else {
					result.degree = format.degree || 0;
				}
				result.start = format.start || 'FFFFFFFF';
				result.end = format.end || 'FFFFFFFF';
			}
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var children = format.fillType === 'pattern' ? [savePatternFill(format)] : [saveGradientFill(format)];

			return toXMLString({
				name: 'fill',
				children: children
			});
		}
	}]);

	return Fills;
}(StylePart);

function savePatternFill(format) {
	var attributes = [['patternType', format.patternType]];
	var children = [toXMLString({
		name: 'fgColor',
		attributes: [['rgb', canonColor(format.fgColor)]]
	}), toXMLString({
		name: 'bgColor',
		attributes: [['rgb', canonColor(format.bgColor)]]
	})];

	return toXMLString({
		name: 'patternFill',
		attributes: attributes,
		children: children
	});
}

function saveGradientFill(format) {
	var attributes = [];

	if (format.degree) {
		attributes.push(['degree', format.degree]);
	} else if (format.left) {
		attributes.push(['type', 'path']);
		attributes.push(['left', format.left]);
		attributes.push(['right', format.right]);
		attributes.push(['top', format.top]);
		attributes.push(['bottom', format.bottom]);
	}

	var children = [toXMLString({
		name: 'stop',
		attributes: [['position', 0]],
		children: [saveColor(format.start)]
	}), toXMLString({
		name: 'stop',
		attributes: [['position', 1]],
		children: [saveColor(format.end)]
	})];

	return toXMLString({
		name: 'gradientFill',
		attributes: attributes,
		children: children
	});
}

module.exports = Fills;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./stylePart":21,"./utils":24}],17:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');

var _require = require('./utils'),
    saveColor = _require.saveColor;

var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fonts.aspx

var Fonts = function (_StylePart) {
	_inherits(Fonts, _StylePart);

	function Fonts(styles) {
		_classCallCheck(this, Fonts);

		var _this = _possibleConstructorReturn(this, (Fonts.__proto__ || Object.getPrototypeOf(Fonts)).call(this, styles, 'fonts', 'font'));

		_this.init();
		_this.lastId = _this.formats.length;
		return _this;
	}

	_createClass(Fonts, [{
		key: 'init',
		value: function init() {
			this.formats.push({ format: this.canon({}) });
		}
	}, {
		key: 'canon',
		value: function canon(format) {
			return Fonts.canon(format);
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			var result = _.assign(formatTo, formatFrom);

			result.color = formatFrom && formatFrom.color || formatTo && formatTo.color;
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			return Fonts.saveFormat(format, styleFormat);
		}
	}], [{
		key: 'canon',
		value: function canon(format) {
			var result = {};

			if (_.has(format, 'bold')) {
				result.bold = !!format.bold;
			}
			if (_.has(format, 'italic')) {
				result.italic = !!format.italic;
			}
			if (format.superscript) {
				result.vertAlign = 'superscript';
			}
			if (format.subscript) {
				result.vertAlign = 'subscript';
			}
			if (format.underline) {
				if (_.indexOf(['double', 'singleAccounting', 'doubleAccounting'], format.underline) !== -1) {
					result.underline = format.underline;
				} else {
					result.underline = true;
				}
			}
			if (_.has(format, 'strike')) {
				result.strike = !!format.strike;
			}
			if (format.outline) {
				result.outline = true;
			}
			if (format.shadow) {
				result.shadow = true;
			}
			if (format.size) {
				result.size = format.size;
			}
			if (format.color) {
				result.color = format.color;
			}
			if (format.fontName) {
				result.fontName = format.fontName;
			}
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var children = [];

			if (format.size) {
				children.push(toXMLString({
					name: 'sz',
					attributes: [['val', format.size]]
				}));
			}
			if (format.fontName) {
				children.push(toXMLString({
					name: 'name',
					attributes: [['val', format.fontName]]
				}));
			}
			if (_.has(format, 'bold')) {
				if (format.bold) {
					children.push(toXMLString({ name: 'b' }));
				} else {
					children.push(toXMLString({ name: 'b', attributes: [['val', 0]] }));
				}
			}
			if (_.has(format, 'italic')) {
				if (format.italic) {
					children.push(toXMLString({ name: 'i' }));
				} else {
					children.push(toXMLString({ name: 'i', attributes: [['val', 0]] }));
				}
			}
			if (format.vertAlign) {
				children.push(toXMLString({
					name: 'vertAlign',
					attributes: [['val', format.vertAlign]]
				}));
			}
			if (format.underline) {
				var attrs = null;

				if (format.underline !== true) {
					attrs = [['val', format.underline]];
				}
				children.push(toXMLString({
					name: 'u',
					attributes: attrs
				}));
			}
			if (_.has(format, 'strike')) {
				if (format.strike) {
					children.push(toXMLString({ name: 'strike' }));
				} else {
					children.push(toXMLString({ name: 'strike', attributes: [['val', 0]] }));
				}
			}
			if (format.shadow) {
				children.push(toXMLString({ name: 'shadow' }));
			}
			if (format.outline) {
				children.push(toXMLString({ name: 'outline' }));
			}
			if (format.color) {
				children.push(saveColor(format.color));
			}

			return toXMLString({
				name: 'font',
				children: children
			});
		}
	}]);

	return Fonts;
}(StylePart);

module.exports = Fonts;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./stylePart":21,"./utils":24}],18:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var NumberFormats = require('./numberFormats');
var Fonts = require('./fonts');
var Fills = require('./fills');
var Borders = require('./borders');
var Cells = require('./cells');
var Tables = require('./tables');
var TableElements = require('./tableElements');
var toXMLString = require('../XMLString');

var Styles = function () {
	function Styles(common) {
		_classCallCheck(this, Styles);

		this.objectId = common.uniqueId('Styles');
		this.numberFormats = new NumberFormats(this);
		this.fonts = new Fonts(this);
		this.fills = new Fills(this);
		this.borders = new Borders(this);
		this.cells = new Cells(this);
		this.tableElements = new TableElements(this);
		this.tables = new Tables(this);
		this.defaultTableStyle = '';
	}

	_createClass(Styles, [{
		key: 'addFormat',
		value: function addFormat(format, name) {
			return this.cells.add(format, name);
		}
	}, {
		key: '_get',
		value: function _get(name) {
			return this.cells.get(name);
		}
	}, {
		key: '_getId',
		value: function _getId(name) {
			return this.cells.getId(name);
		}
	}, {
		key: '_addInvisibleFormat',
		value: function _addInvisibleFormat(format) {
			var style = this.cells.cutVisible(this.cells.fullGet(format));

			if (!_.isEmpty(style)) {
				return this.addFormat(style);
			}
		}
	}, {
		key: '_merge',
		value: function _merge(columnFormat, rowFormat, cellFormat) {
			var count = Number(Boolean(columnFormat)) + Number(Boolean(rowFormat)) + Number(Boolean(cellFormat));

			if (count === 0) {
				return null;
			} else if (count === 1) {
				return this.cells.add(columnFormat || rowFormat || cellFormat);
			} else {
				var format = {};

				if (columnFormat) {
					format = this.cells.merge(format, this.cells.fullGet(columnFormat));
				}
				if (rowFormat) {
					format = this.cells.merge(format, this.cells.fullGet(rowFormat));
				}
				if (cellFormat) {
					format = this.cells.merge(format, this.cells.fullGet(cellFormat));
				}
				return this.cells.add(format, null, { merge: true });
			}
		}
	}, {
		key: 'addFontFormat',
		value: function addFontFormat(format, name) {
			return this.fonts.add(format, name);
		}
	}, {
		key: 'addBorderFormat',
		value: function addBorderFormat(format, name) {
			return this.borders.add(format, name);
		}
	}, {
		key: 'addPatternFormat',
		value: function addPatternFormat(format, name) {
			return this.fills.add(format, name, { fillType: 'pattern' });
		}
	}, {
		key: 'addGradientFormat',
		value: function addGradientFormat(format, name) {
			return this.fills.add(format, name, { fillType: 'gradient' });
		}
	}, {
		key: 'addNumberFormat',
		value: function addNumberFormat(format, name) {
			return this.numberFormats.add(format, name);
		}
	}, {
		key: 'addTableFormat',
		value: function addTableFormat(format, name) {
			return this.tables.add(format, name);
		}
	}, {
		key: 'addTableElementFormat',
		value: function addTableElementFormat(format, name) {
			return this.tableElements.add(format, name);
		}
	}, {
		key: 'setDefaultTableStyle',
		value: function setDefaultTableStyle(name) {
			this.tables.defaultTableStyle = name;
		}
	}, {
		key: 'save',
		value: function save() {
			return toXMLString({
				name: 'styleSheet',
				ns: 'spreadsheetml',
				children: [this.numberFormats.save(), this.fonts.save(), this.fills.save(), this.borders.save(), this.cells.save(), this.tableElements.save(), this.tables.save()]
			});
		}
	}]);

	return Styles;
}();

module.exports = Styles;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./borders":14,"./cells":15,"./fills":16,"./fonts":17,"./numberFormats":19,"./tableElements":22,"./tables":23}],19:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var toXMLString = require('../XMLString');

var PREDEFINED = {
	date: 14, //mm-dd-yy
	time: 21 //h:mm:ss
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats.aspx

var NumberFormats = function (_StylePart) {
	_inherits(NumberFormats, _StylePart);

	function NumberFormats(styles) {
		_classCallCheck(this, NumberFormats);

		var _this = _possibleConstructorReturn(this, (NumberFormats.__proto__ || Object.getPrototypeOf(NumberFormats)).call(this, styles, 'numFmts', 'numberFormat'));

		_this.init();
		_this.lastId = 164;
		return _this;
	}

	_createClass(NumberFormats, [{
		key: 'init',
		value: function init() {
			var _this2 = this;

			_.forEach(PREDEFINED, function (formatId, format) {
				_this2.formatsByNames[format] = {
					formatId: formatId,
					format: format
				};
			});
		}
	}, {
		key: 'canon',
		value: function canon(format) {
			return NumberFormats.canon(format);
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			return formatFrom || formatTo;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			return NumberFormats.saveFormat(format, styleFormat);
		}
	}], [{
		key: 'canon',
		value: function canon(format) {
			return format;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			var attributes = [['numFmtId', styleFormat.formatId], ['formatCode', format]];

			return toXMLString({
				name: 'numFmt',
				attributes: attributes
			});
		}
	}]);

	return NumberFormats;
}(StylePart);

module.exports = NumberFormats;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./stylePart":21}],20:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.protection.aspx

function canon(format) {
	var result = {};

	if (_.has(format, 'locked')) {
		result.locked = format.locked ? 1 : 0;
	}
	if (_.has(format, 'hidden')) {
		result.hidden = format.hidden ? 1 : 0;
	}

	return _.isEmpty(result) ? null : result;
}

function merge(formatTo, formatFrom) {
	return _.assign(formatTo, formatFrom);
}

function saveFormat(format) {
	return toXMLString({
		name: 'protection',
		attributes: _.toPairs(format)
	});
}

module.exports = {
	canon: canon,
	merge: merge,
	saveFormat: saveFormat
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],21:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

var StylePart = function () {
	function StylePart(styles, saveName, formatName) {
		_classCallCheck(this, StylePart);

		this.styles = styles;
		this.saveName = saveName;
		this.formatName = formatName;
		this.lastName = 1;
		this.lastId = 0;
		this.saveEmpty = true;
		this.formats = [];
		this.formatsByData = Object.create(null);
		this.formatsByNames = Object.create(null);
	}

	_createClass(StylePart, [{
		key: 'add',
		value: function add(format, name, flags) {
			if (name && this.formatsByNames[name]) {
				var _canonFormat = this.canon(format, flags);
				var _stringFormat = _.isObject(_canonFormat) ? JSON.stringify(_canonFormat) : _canonFormat;

				if (_stringFormat !== this.formatsByNames[name].stringFormat) {
					this._add(_canonFormat, _stringFormat, name);
				}
				return name;
			}

			//first argument is format name
			if (!name && _.isString(format) && this.formatsByNames[format]) {
				return format;
			}

			var canonFormat = this.canon(format, flags);
			var stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;
			var styleFormat = this.formatsByData[stringFormat];

			if (!styleFormat) {
				styleFormat = this._add(canonFormat, stringFormat, name);
			} else if (name && !this.formatsByNames[name]) {
				styleFormat.name = name;
				this.formatsByNames[name] = styleFormat;
			}
			return styleFormat.name;
		}
	}, {
		key: '_add',
		value: function _add(canonFormat, stringFormat, name) {
			name = name || this.formatName + this.lastName++;

			var styleFormat = {
				name: name,
				formatId: this.lastId++,
				format: canonFormat,
				stringFormat: stringFormat
			};

			this.formats.push(styleFormat);
			this.formatsByData[stringFormat] = styleFormat;
			this.formatsByNames[name] = styleFormat;

			return styleFormat;
		}
	}, {
		key: 'canon',
		value: function canon(format) {
			return format;
		}
	}, {
		key: 'get',
		value: function get(format) {
			if (_.isString(format)) {
				var styleFormat = this.formatsByNames[format];

				return styleFormat ? styleFormat.format : format;
			}
			return format;
		}
	}, {
		key: 'getId',
		value: function getId(name) {
			var styleFormat = this.formatsByNames[name];

			return styleFormat ? styleFormat.formatId : null;
		}
	}, {
		key: 'save',
		value: function save() {
			var _this = this;

			if (this.saveEmpty !== false || this.formats.length) {
				var attributes = [['count', this.formats.length]];
				var children = this.formats.map(function (format) {
					return _this.saveFormat(format.format, format);
				});

				this.saveCollectionExt(attributes, children);

				return toXMLString({
					name: this.saveName,
					attributes: attributes,
					children: children
				});
			}
			return '';
		}
	}, {
		key: 'saveCollectionExt',
		value: function saveCollectionExt() {}
	}, {
		key: 'saveFormat',
		value: function saveFormat() {
			return '';
		}
	}]);

	return StylePart;
}();

module.exports = StylePart;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],22:[function(require,module,exports){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var StylePart = require('./stylePart');
var NumberFormats = require('./numberFormats');
var Fonts = require('./fonts');
var Fills = require('./fills');
var Borders = require('./borders');
var alignment = require('./alignment');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.differentialformats.aspx

var TableElements = function (_StylePart) {
	_inherits(TableElements, _StylePart);

	function TableElements(styles) {
		_classCallCheck(this, TableElements);

		return _possibleConstructorReturn(this, (TableElements.__proto__ || Object.getPrototypeOf(TableElements)).call(this, styles, 'dxfs', 'tableElement'));
	}

	_createClass(TableElements, [{
		key: 'canon',
		value: function canon(format) {
			var result = {};

			if (format.format) {
				result.format = NumberFormats.canon(format.format);
			}
			if (format.font) {
				result.font = Fonts.canon(format.font);
			}
			if (format.pattern) {
				result.fill = Fills.canon(format.pattern, { fillType: 'pattern', isTable: true });
			} else if (format.gradient) {
				result.fill = Fills.canon(format.gradient, { fillType: 'gradient' });
			}
			if (format.border) {
				result.border = Borders.canon(format.border);
			}
			result.alignment = alignment.canon(format);
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var children = [];

			if (format.font) {
				children.push(Fonts.saveFormat(format.font));
			}
			if (format.fill) {
				children.push(Fills.saveFormat(format.fill));
			}
			if (format.border) {
				children.push(Borders.saveFormat(format.border));
			}
			if (format.format) {
				children.push(NumberFormats.saveFormat(format.format));
			}
			if (format.alignment && format.alignment.length) {
				children.push(alignment.save(format.alignment));
			}

			return toXMLString({
				name: 'dxf',
				children: children
			});
		}
	}]);

	return TableElements;
}(StylePart);

module.exports = TableElements;

},{"../XMLString":2,"./alignment":13,"./borders":14,"./fills":16,"./fonts":17,"./numberFormats":19,"./stylePart":21}],23:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestylevalues.aspx
var ELEMENTS = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn', 'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe', 'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell'];
var SIZED_ELEMENTS = ['firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyles.aspx

var Tables = function (_StylePart) {
	_inherits(Tables, _StylePart);

	function Tables(styles) {
		_classCallCheck(this, Tables);

		var _this = _possibleConstructorReturn(this, (Tables.__proto__ || Object.getPrototypeOf(Tables)).call(this, styles, 'tableStyles', 'table'));

		_this.saveEmpty = false;
		return _this;
	}

	_createClass(Tables, [{
		key: 'canon',
		value: function canon(format) {
			var result = {};
			var styles = this.styles;

			_.forEach(format, function (value, key) {
				if (_.includes(ELEMENTS, key)) {
					var style = void 0;
					var size = null;

					if (value.style) {
						style = styles.tableElements.add(value.style);
						if (value.size > 1 && _.includes(SIZED_ELEMENTS, key)) {
							size = value.size;
						}
					} else {
						style = styles.tableElements.add(value);
					}
					result[key] = {
						style: style,
						size: size
					};
				}
			});

			return result;
		}
	}, {
		key: 'saveCollectionExt',
		value: function saveCollectionExt(attributes) {
			if (this.styles.defaultTableStyle) {
				attributes.push(['defaultTableStyle', this.styles.defaultTableStyle]);
			}
		}
		//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyleelement.aspx

	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			var styles = this.styles;
			var attributes = [['name', styleFormat.name], ['pivot', 0]];
			var children = [];

			_.forEach(format, function (value, key) {
				var style = value.style;
				var size = value.size;
				var attributes = [['type', key], ['dxfId', styles.tableElements.getId(style)]];

				if (size) {
					attributes.push(['size', size]);
				}

				children.push(toXMLString({
					name: 'tableStyleElement',
					attributes: attributes
				}));
			});
			attributes.push(['count', children.length]);

			return toXMLString({
				name: 'tableStyle',
				attributes: attributes,
				children: children
			});
		}
	}]);

	return Tables;
}(StylePart);

module.exports = Tables;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./stylePart":21}],24:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

function canonColor(color) {
	return color[0] === '#' ? 'FF' + color.substr(1) : color;
}

function saveColor(color) {
	if (_.isString(color)) {
		return toXMLString({
			name: 'color',
			attributes: [['rgb', canonColor(color)]]
		});
	} else {
		var attributes = [];
		if (!_.isUndefined(color.tint)) {
			attributes.push(['tint', color.tint]);
		}
		if (!_.isUndefined(color.auto)) {
			attributes.push(['auto', !!color.auto]);
		}
		if (!_.isUndefined(color.theme)) {
			attributes.push(['theme', color.theme]);
		}

		return toXMLString({
			name: 'color',
			attributes: attributes
		});
	}
}

module.exports = {
	canonColor: canonColor,
	saveColor: saveColor
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],25:[function(require,module,exports){
(function (global){
'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');
var toXMLString = require('./XMLString');

var Table = function () {
	function Table(worksheet, config) {
		_classCallCheck(this, Table);

		this.worksheet = worksheet;
		this.common = worksheet.common;

		this.tableId = this.common.uniqueIdForSpace('Table');
		this.objectId = 'Table' + this.tableId;
		this.name = this.objectId;
		this.displayName = this.objectId;
		this.headerRowCount = 1;
		this.beginCell = null;
		this.endCell = null;
		this.totalsRowCount = 0;
		this.totalRow = null;
		this.themeStyle = null;

		_.extend(this, config);
	}

	_createClass(Table, [{
		key: 'end',
		value: function end() {
			return this.worksheet;
		}
	}, {
		key: 'setReferenceRange',
		value: function setReferenceRange(beginCell, endCell) {
			this.beginCell = util.canonCell(beginCell);
			this.endCell = util.canonCell(endCell);
			return this;
		}
	}, {
		key: 'addTotalRow',
		value: function addTotalRow(totalRow) {
			this.totalRow = totalRow;
			this.totalsRowCount = 1;
			return this;
		}
	}, {
		key: 'setTheme',
		value: function setTheme(theme) {
			this.themeStyle = theme;
			return this;
		}
		/**
   * Expects an object with the following properties:
   * caseSensitive (boolean)
   * dataRange
   * columnSort (assumes true)
   * sortDirection
   * sortRange (defaults to dataRange)
   */

	}, {
		key: 'setSortState',
		value: function setSortState(state) {
			this.sortState = state;
			return this;
		}
	}, {
		key: '_prepare',
		value: function _prepare(worksheetData) {
			var _this = this;

			var SUB_TOTAL_FUNCTIONS = ['average', 'countNums', 'count', 'max', 'min', 'stdDev', 'sum', 'var'];
			var SUB_TOTAL_NUMS = [101, 102, 103, 104, 105, 107, 109, 110];

			if (this.totalRow) {
				(function () {
					var tableName = _this.name;
					var beginCell = util.letterToPosition(_this.beginCell);
					var endCell = util.letterToPosition(_this.endCell);
					var firstRow = beginCell.y - 1;
					var firstColumn = beginCell.x - 1;
					var lastRow = endCell.y - 1;
					var headerRow = worksheetData[firstRow] || [];
					var totalRow = worksheetData[lastRow + 1];

					if (!totalRow) {
						totalRow = [];
						worksheetData[lastRow + 1] = totalRow;
					}

					_.forEach(_this.totalRow, function (cell, i) {
						var headerValue = headerRow[firstColumn + i];
						var funcIndex = void 0;

						if ((typeof headerValue === 'undefined' ? 'undefined' : _typeof(headerValue)) === 'object') {
							headerValue = headerValue.value;
						}
						cell.name = headerValue;
						if (cell.totalsRowLabel) {
							totalRow[firstColumn + i] = {
								value: cell.totalsRowLabel,
								type: 'string'
							};
						} else if (cell.totalsRowFunction) {
							funcIndex = _.indexOf(SUB_TOTAL_FUNCTIONS, cell.totalsRowFunction);

							if (funcIndex !== -1) {
								totalRow[firstColumn + i] = {
									value: 'SUBTOTAL(' + SUB_TOTAL_NUMS[funcIndex] + ',' + tableName + '[' + headerValue + '])',
									type: 'formula'
								};
							}
						}
					});
				})();
			}
		}
		//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.table.aspx

	}, {
		key: '_save',
		value: function _save() {
			var attributes = [['id', this.tableId], ['name', this.name], ['displayName', this.displayName]];
			var children = [];
			var end = util.letterToPosition(this.endCell);
			var ref = this.beginCell + ':' + util.positionToLetter(end.x, end.y + this.totalsRowCount);

			attributes.push(['ref', ref]);
			attributes.push(['totalsRowCount', this.totalsRowCount]);
			attributes.push(['headerRowCount', this.headerRowCount]);

			children.push(saveAutoFilter(this.beginCell, this.endCell));
			children.push(saveTableColumns(this.totalRow));
			children.push(saveTableStyleInfo(this.common, this.themeStyle));

			return toXMLString({
				name: 'table',
				ns: 'spreadsheetml',
				attributes: attributes,
				children: children
			});
		}
	}]);

	return Table;
}();

function saveAutoFilter(beginCell, endCell) {
	return toXMLString({
		name: 'autoFilter',
		attributes: ['ref', beginCell + ':' + endCell]
	});
}

function saveTableColumns(totalRow) {
	var attributes = [['count', totalRow.length]];
	var children = _.map(totalRow, function (cell, index) {
		var attributes = [['id', index + 1], ['name', cell.name]];

		if (cell.totalsRowFunction) {
			attributes.push(['totalsRowFunction', cell.totalsRowFunction]);
		}
		if (cell.totalsRowLabel) {
			attributes.push(['totalsRowLabel', cell.totalsRowLabel]);
		}

		return toXMLString({
			name: 'tableColumn',
			attributes: attributes
		});
	});

	return toXMLString({
		name: 'tableColumns',
		attributes: attributes,
		children: children
	});
}

function saveTableStyleInfo(common, themeStyle) {
	var attributes = [['name', themeStyle]];
	var format = common.styles.tables.get(themeStyle);
	var isRowStripes = false;
	var isColumnStripes = false;
	var isFirstColumn = false;
	var isLastColumn = false;

	if (format) {
		isRowStripes = format.firstRowStripe || format.secondRowStripe;
		isColumnStripes = format.firstColumnStripe || format.secondColumnStripe;
		isFirstColumn = format.firstColumn;
		isLastColumn = format.lastColumn;
	}
	attributes.push(['showRowStripes', isRowStripes ? '1' : '0'], ['showColumnStripes', isColumnStripes ? '1' : '0'], ['showFirstColumn', isFirstColumn ? '1' : '0'], ['showLastColumn', isLastColumn ? '1' : '0']);

	return toXMLString({
		name: 'tableStyleInfo',
		attributes: attributes
	});
}

module.exports = Table;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./util":26}],26:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var LETTER_REFS = {};

function positionToLetter(x, y) {
	var result = LETTER_REFS[x];

	if (!result) {
		var string = '';
		var num = x;
		var index = void 0;

		do {
			index = (num - 1) % 26;
			string = alphabet[index] + string;
			num = (num - (index + 1)) / 26;
		} while (num > 0);

		LETTER_REFS[x] = string;
		result = string;
	}
	return result + (y || '');
}

function letterToPosition(cell) {
	var x = 0;
	var y = 0;

	for (var i = 0, len = cell.length; i < len; i++) {
		var charCode = cell.charCodeAt(i);
		if (charCode >= 65) {
			x = x * 26 + charCode - 64;
		} else {
			y = parseInt(cell.slice(i), 10);
			break;
		}
	}
	return {
		x: x || 1,
		y: y || 1
	};
}

module.exports = {
	pixelsToEMUs: function pixelsToEMUs(pixels) {
		return Math.round(pixels * 914400 / 96);
	},
	canonCell: function canonCell(cell) {
		if (_.isObject(cell)) {
			return positionToLetter(cell.c || 1, cell.r || 1);
		}
		return cell;
	},


	positionToLetter: positionToLetter,
	letterToPosition: letterToPosition,

	xmlPrefix: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',

	schemas: {
		'worksheet': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
		'sharedStrings': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
		'stylesheet': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
		'relationships': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
		'relationshipPackage': 'http://schemas.openxmlformats.org/package/2006/relationships',
		'contentTypes': 'http://schemas.openxmlformats.org/package/2006/content-types',
		'spreadsheetml': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
		'markupCompat': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
		'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
		'officeDocument': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
		'package': 'http://schemas.openxmlformats.org/package/2006/relationships',
		'table': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
		'spreadsheetDrawing': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
		'drawing': 'http://schemas.openxmlformats.org/drawingml/2006/main',
		'drawingRelationship': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
		'image': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
		'chart': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
		'hyperlink': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
	}
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{}],27:[function(require,module,exports){
(function (global){
'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');
var Common = require('./common');
var Relations = require('./relations');
var Worksheet = require('./worksheet');
var toXMLString = require('./XMLString');

var Workbook = function () {
	function Workbook() {
		_classCallCheck(this, Workbook);

		this.common = new Common();
		this.styles = this.common.styles;
		this.images = this.common.images;

		this.objectId = this.common.uniqueId('Workbook');

		this.relations = new Relations(this.common);
		this.relations.addRelation(this.styles, 'stylesheet');
	}

	_createClass(Workbook, [{
		key: 'addWorksheet',
		value: function addWorksheet(config) {
			config = _.defaults(config, {
				name: this.common.getNewWorksheetDefaultName()
			});
			var worksheet = new Worksheet(this, config);
			this.common.addWorksheet(worksheet);
			this.relations.addRelation(worksheet, 'worksheet');

			return worksheet;
		}
	}, {
		key: 'addFormat',
		value: function addFormat(format, name) {
			return this.styles.addFormat(format, name);
		}
	}, {
		key: 'addFontFormat',
		value: function addFontFormat(format, name) {
			return this.styles.addFontFormat(format, name);
		}
	}, {
		key: 'addBorderFormat',
		value: function addBorderFormat(format, name) {
			return this.styles.addBorderFormat(format, name);
		}
	}, {
		key: 'addPatternFormat',
		value: function addPatternFormat(format, name) {
			return this.styles.addPatternFormat(format, name);
		}
	}, {
		key: 'addGradientFormat',
		value: function addGradientFormat(format, name) {
			return this.styles.addGradientFormat(format, name);
		}
	}, {
		key: 'addNumberFormat',
		value: function addNumberFormat(format, name) {
			return this.styles.addNumberFormat(format, name);
		}
	}, {
		key: 'addTableFormat',
		value: function addTableFormat(format, name) {
			return this.styles.addTableFormat(format, name);
		}
	}, {
		key: 'addTableElementFormat',
		value: function addTableElementFormat(format, name) {
			return this.styles.addTableElementFormat(format, name);
		}
	}, {
		key: 'setDefaultTableStyle',
		value: function setDefaultTableStyle(name) {
			this.styles.setDefaultTableStyle(name);
			return this;
		}
	}, {
		key: 'addImage',
		value: function addImage(data, type, name) {
			return this.images.addImage(data, type, name);
		}
	}, {
		key: '_generateFiles',
		value: function _generateFiles(zip, canStream) {
			prepareWorksheets(this.common);

			saveWorksheets(zip, canStream, this.common);
			saveTables(zip, this.common);
			saveImages(zip, this.images);
			saveDrawings(zip, this.common);
			saveStyles(zip, this.styles);
			saveStrings(zip, canStream, this.relations, this.common);
			zip.file('[Content_Types].xml', createContentTypes(this.common));
			zip.file('_rels/.rels', createWorkbookRelationship());
			zip.file('xl/workbook.xml', this._save());
			zip.file('xl/_rels/workbook.xml.rels', this.relations.save());
		}
	}, {
		key: '_save',
		value: function _save() {
			return toXMLString({
				name: 'workbook',
				ns: 'spreadsheetml',
				attributes: [['xmlns:r', util.schemas.relationships]],
				children: [bookViewsXML(this.common), sheetsXML(this.relations, this.common), definedNamesXML(this.common)]
			});
		}
	}]);

	return Workbook;
}();

function bookViewsXML(common) {
	var activeTab = 0;

	if (common.activeWorksheet) {
		(function () {
			var activeWorksheetId = common.activeWorksheet.objectId;

			activeTab = Math.max(activeTab, _.findIndex(common.worksheets, function (worksheet) {
				return worksheet.objectId === activeWorksheetId;
			}));
		})();
	}

	return toXMLString({
		name: 'bookViews',
		children: [toXMLString({
			name: 'workbookView',
			attributes: [['activeTab', activeTab]]
		})]
	});
}

function sheetsXML(relations, common) {
	var maxWorksheetNameLength = 31;
	var children = _.map(common.worksheets, function (worksheet, index) {
		// Microsoft Excel (2007, 2013) do not allow worksheet names longer than 31 characters
		// if the worksheet name is longer, Excel displays an 'Excel found unreadable content...' popup when opening the file
		if (worksheet.name.length > maxWorksheetNameLength) {
			throw 'Microsoft Excel requires work sheet names to be less than ' + (maxWorksheetNameLength + 1) + ' characters long, work sheet name "' + worksheet.name + '" is ' + worksheet.name.length + ' characters long';
		}

		return toXMLString({
			name: 'sheet',
			attributes: [['name', worksheet.name], ['sheetId', index + 1], ['r:id', relations.getRelationshipId(worksheet)], ['state', worksheet.getState()]]
		});
	});

	return toXMLString({
		name: 'sheets',
		children: children
	});
}

function definedNamesXML(common) {
	var isPrintTitles = _.some(common.worksheets, function (worksheet) {
		return worksheet._printTitles && (worksheet._printTitles.topTo >= 0 || worksheet._printTitles.leftTo >= 0);
	});

	if (isPrintTitles) {
		var _ret2 = function () {
			var children = [];

			_.forEach(common.worksheets, function (worksheet, index) {
				var entry = worksheet._printTitles;

				if (entry && (entry.topTo >= 0 || entry.leftTo >= 0)) {
					var name = worksheet.name;
					var value = '';

					if (entry.topTo >= 0) {
						value = name + '!$' + (entry.topFrom + 1) + ':$' + (entry.topTo + 1);
						if (entry.leftTo >= 0) {
							value += ',';
						}
					}
					if (entry.leftTo >= 0) {
						value += name + '!$' + util.positionToLetter(entry.leftFrom + 1) + ':$' + util.positionToLetter(entry.leftTo + 1);
					}

					children.push(toXMLString({
						name: 'definedName',
						value: value,
						attributes: [['name', '_xlnm.Print_Titles'], ['localSheetId', index]]
					}));
				}
			});

			return {
				v: toXMLString({
					name: 'definedNames',
					children: children
				})
			};
		}();

		if ((typeof _ret2 === 'undefined' ? 'undefined' : _typeof(_ret2)) === "object") return _ret2.v;
	}
	return '';
}

function prepareWorksheets(common) {
	_.forEach(common.worksheets, function (worksheet) {
		worksheet._prepare();
	});
}

function saveWorksheets(zip, canStream, common) {
	_.forEach(common.worksheets, function (worksheet) {
		zip.file(worksheet.path, worksheet._save(canStream));
		zip.file(worksheet.relationsPath, worksheet.relations.save());
	});
}

function saveTables(zip, common) {
	_.forEach(common.tables, function (table) {
		zip.file(table.path, table._save());
	});
}

function saveImages(zip, images) {
	_.forEach(images.getImages(), function (image) {
		zip.file(image.path, image.data, { base64: true, binary: true });
		image.data = null;
	});
	images.removeImages();
}

function saveDrawings(zip, common) {
	_.forEach(common.drawings, function (drawing) {
		zip.file(drawing.path, drawing.save());
		zip.file(drawing.relationsPath, drawing.relations.save());
	});
}

function saveStyles(zip, styles) {
	zip.file('xl/styles.xml', styles.save());
}

function saveStrings(zip, canStream, relations, common) {
	if (common.strings.isStrings()) {
		relations.addRelation(common.strings, 'sharedStrings');
		zip.file('xl/sharedStrings.xml', common.strings.save(canStream));
	}
}

function createContentTypes(common) {
	var children = [];

	children.push(toXMLString({
		name: 'Default',
		attributes: [['Extension', 'rels'], ['ContentType', 'application/vnd.openxmlformats-package.relationships+xml']]
	}));
	children.push(toXMLString({
		name: 'Default',
		attributes: [['Extension', 'xml'], ['ContentType', 'application/xml']]
	}));
	children.push(toXMLString({
		name: 'Override',
		attributes: [['PartName', '/xl/workbook.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml']]
	}));
	if (common.strings.isStrings()) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/sharedStrings.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml']]
		}));
	}
	children.push(toXMLString({
		name: 'Override',
		attributes: [['PartName', '/xl/styles.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml']]
	}));

	_.forEach(common.worksheets, function (worksheet, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/worksheets/sheet' + (index + 1) + '.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']]
		}));
	});
	_.forEach(common.tables, function (table, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/tables/table' + (index + 1) + '.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']]
		}));
	});
	_.forEach(common.images.getExtensions(), function (contentType, extension) {
		children.push(toXMLString({
			name: 'Default',
			attributes: [['Extension', extension], ['ContentType', contentType]]
		}));
	});
	_.forEach(common.drawings, function (drawing, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/drawings/drawing' + (index + 1) + '.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml']]
		}));
	});

	return toXMLString({
		name: 'Types',
		ns: 'contentTypes',
		children: children
	});
}

function createWorkbookRelationship() {
	return toXMLString({
		name: 'Relationships',
		ns: 'relationshipPackage',
		children: [toXMLString({
			name: 'Relationship',
			attributes: [['Id', 'rId1'], ['Type', util.schemas.officeDocument], ['Target', 'xl/workbook.xml']]
		})]
	});
}

module.exports = Workbook;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./common":4,"./relations":12,"./util":26,"./worksheet":30}],28:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var Tables = require('./tables');
var Drawings = require('../drawings');
var toXMLString = require('../XMLString');

var DrawingsExt = function (_Tables) {
	_inherits(DrawingsExt, _Tables);

	function DrawingsExt() {
		_classCallCheck(this, DrawingsExt);

		var _this = _possibleConstructorReturn(this, (DrawingsExt.__proto__ || Object.getPrototypeOf(DrawingsExt)).call(this));

		_this._drawings = null;
		return _this;
	}

	_createClass(DrawingsExt, [{
		key: 'setImage',
		value: function setImage(image, config) {
			return this._setDrawing(image, config, 'anchor');
		}
	}, {
		key: 'setImageOneCell',
		value: function setImageOneCell(image, config) {
			return this._setDrawing(image, config, 'oneCell');
		}
	}, {
		key: 'setImageAbsolute',
		value: function setImageAbsolute(image, config) {
			return this._setDrawing(image, config, 'absolute');
		}
	}, {
		key: '_setDrawing',
		value: function _setDrawing(image, config, anchorType) {
			var name = void 0;

			if (!this._drawings) {
				this._drawings = new Drawings(this.common);

				this.common.addDrawings(this._drawings);
				this.relations.addRelation(this._drawings, 'drawingRelationship');
			}

			if (_.isObject(image)) {
				name = this.common.images.addImage(image.data, image.type);
			} else {
				name = image;
			}

			this._drawings.addImage(name, config, anchorType);
			return this;
		}
	}, {
		key: '_insertDrawing',
		value: function _insertDrawing(colIndex, rowIndex, image) {
			if (typeof image === 'string' || image.data) {
				this._setDrawing(image, { c: colIndex + 1, r: rowIndex + 1 }, 'anchor');
			} else {
				var config = image.config || {};
				config.cell = { c: colIndex + 1, r: rowIndex + 1 };

				this._setDrawing(image.image, config, 'anchor');
			}
		}
	}, {
		key: '_saveDrawing',
		value: function _saveDrawing() {
			if (this._drawings) {
				return toXMLString({
					name: 'drawing',
					attributes: [['r:id', this.relations.getRelationshipId(this._drawings)]]
				});
			}
			return '';
		}
	}]);

	return DrawingsExt;
}(Tables);

module.exports = DrawingsExt;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../drawings":9,"./tables":36}],29:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var MergedCells = require('./mergedCells');
var toXMLString = require('../XMLString');

var Hyperlinks = function (_MergedCells) {
	_inherits(Hyperlinks, _MergedCells);

	function Hyperlinks() {
		_classCallCheck(this, Hyperlinks);

		var _this = _possibleConstructorReturn(this, (Hyperlinks.__proto__ || Object.getPrototypeOf(Hyperlinks)).call(this));

		_this._hyperlinks = [];
		return _this;
	}

	_createClass(Hyperlinks, [{
		key: 'setHyperlink',
		value: function setHyperlink(hyperlink) {
			hyperlink.objectId = this.common.uniqueId('hyperlink');
			this.relations.addRelation({
				objectId: hyperlink.objectId,
				target: hyperlink.location,
				targetMode: hyperlink.targetMode || 'External'
			}, 'hyperlink');
			this._hyperlinks.push(hyperlink);
			return this;
		}
	}, {
		key: '_insertHyperlink',
		value: function _insertHyperlink(colIndex, rowIndex, hyperlink) {
			var location = void 0;
			var targetMode = void 0;
			var tooltip = void 0;

			if (typeof hyperlink === 'string') {
				location = hyperlink;
			} else {
				location = hyperlink.location;
				targetMode = hyperlink.targetMode;
				tooltip = hyperlink.tooltip;
			}
			this.setHyperlink({
				cell: { c: colIndex + 1, r: rowIndex + 1 },
				location: location,
				targetMode: targetMode,
				tooltip: tooltip
			});
		}
	}, {
		key: '_saveHyperlinks',
		value: function _saveHyperlinks() {
			var _this2 = this;

			if (this._hyperlinks.length > 0) {
				var children = _.map(this._hyperlinks, function (hyperlink) {
					var attributes = [['ref', util.canonCell(hyperlink.cell)], ['r:id', _this2.relations.getRelationshipId(hyperlink)]];

					if (hyperlink.tooltip) {
						attributes.push(['tooltip', hyperlink.tooltip]);
					}
					return toXMLString({
						name: 'hyperlink',
						attributes: attributes
					});
				});

				return toXMLString({
					name: 'hyperlinks',
					children: children
				});
			}
			return '';
		}
	}]);

	return Hyperlinks;
}(MergedCells);

module.exports = Hyperlinks;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":26,"./mergedCells":31}],30:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var WorksheetSave = require('./save');
var Relations = require('../relations');

var Worksheet = function (_WorksheetSave) {
	_inherits(Worksheet, _WorksheetSave);

	function Worksheet(workbook) {
		var config = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};

		_classCallCheck(this, Worksheet);

		var _this = _possibleConstructorReturn(this, (Worksheet.__proto__ || Object.getPrototypeOf(Worksheet)).call(this, workbook, config));

		_this.workbook = workbook;
		_this.common = workbook.common;

		_this.objectId = _this.common.uniqueId('Worksheet');
		_this.data = [];
		_this.columns = [];
		_this.rows = [];

		_this.name = config.name;
		_this.state = config.state || 'visible';
		_this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
		_this.relations = new Relations(_this.common);
		return _this;
	}

	_createClass(Worksheet, [{
		key: 'end',
		value: function end() {
			return this.workbook;
		}
	}, {
		key: 'setActive',
		value: function setActive() {
			this.common.setActiveWorksheet(this);
			return this;
		}
	}, {
		key: 'setVisible',
		value: function setVisible() {
			this.setState('visible');
			return this;
		}
	}, {
		key: 'setHidden',
		value: function setHidden() {
			this.setState('hidden');
			return this;
		}
		/**
   * //http://www.datypic.com/sc/ooxml/t-ssml_ST_SheetState.html
   * @param state - visible | hidden | veryHidden
   */

	}, {
		key: 'setState',
		value: function setState(state) {
			this.state = state;
			return this;
		}
	}, {
		key: 'getState',
		value: function getState() {
			return this.state;
		}
	}, {
		key: 'setRows',
		value: function setRows(startRow, rows) {
			var _this2 = this;

			if (!rows) {
				rows = startRow;
				startRow = 0;
			} else {
				--startRow;
			}
			_.forEach(rows, function (row, i) {
				_this2.rows[startRow + i] = row;
			});
			return this;
		}
	}, {
		key: 'setRow',
		value: function setRow(rowIndex, meta) {
			this.rows[--rowIndex] = meta;
			return this;
		}
		/**
   * http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.aspx
   */

	}, {
		key: 'setColumns',
		value: function setColumns(startColumn, columns) {
			var _this3 = this;

			if (!columns) {
				columns = startColumn;
				startColumn = 0;
			} else {
				--startColumn;
			}
			_.forEach(columns, function (column, i) {
				_this3.columns[startColumn + i] = column;
			});
			return this;
		}
	}, {
		key: 'setColumn',
		value: function setColumn(columnIndex, column) {
			this.columns[--columnIndex] = column;
			return this;
		}
	}, {
		key: 'setData',
		value: function setData(offset, data) {
			var _this4 = this;

			var startRow = this.data.length;

			if (!data) {
				data = offset;
			} else {
				startRow += offset;
			}
			_.forEach(data, function (row, i) {
				_this4.data[startRow + i] = row;
			});
			return this;
		}
	}]);

	return Worksheet;
}(WorksheetSave);

module.exports = Worksheet;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../relations":12,"./save":34}],31:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var DrawingsExt = require('./drawing');
var toXMLString = require('../XMLString');

var MergedCells = function (_DrawingsExt) {
	_inherits(MergedCells, _DrawingsExt);

	function MergedCells() {
		_classCallCheck(this, MergedCells);

		var _this = _possibleConstructorReturn(this, (MergedCells.__proto__ || Object.getPrototypeOf(MergedCells)).call(this));

		_this._mergedCells = [];
		return _this;
	}

	_createClass(MergedCells, [{
		key: 'mergeCells',
		value: function mergeCells(cell1, cell2) {
			this._mergedCells.push([cell1, cell2]);
			return this;
		}
	}, {
		key: '_insertMergeCells',
		value: function _insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan) {
			if (colSpan) {
				for (var j = 0; j < colSpan; j++) {
					dataRow.splice(colIndex + 1, 0, { style: null, type: 'empty' });
				}
			}
			if (rowSpan) {
				colSpan += 1;

				for (var i = 0; i < rowSpan; i++) {
					//todo: original data changed
					var row = this.data[rowIndex + i + 1];

					if (!row) {
						row = [];
						this.data[rowIndex + i + 1] = row;
					}

					if (row.length > colIndex) {
						for (var _j = 0; _j < colSpan; _j++) {
							row.splice(colIndex, 0, { style: null, type: 'empty' });
						}
					} else {
						for (var _j2 = 0; _j2 < colSpan; _j2++) {
							row[colIndex + _j2] = { style: null, type: 'empty' };
						}
					}
				}
			}
		}
	}, {
		key: '_saveMergeCells',
		value: function _saveMergeCells() {
			if (this._mergedCells.length > 0) {
				var children = _.map(this._mergedCells, function (mergeCell) {
					return toXMLString({
						name: 'mergeCell',
						attributes: [['ref', util.canonCell(mergeCell[0]) + ':' + util.canonCell(mergeCell[1])]]
					});
				});

				return toXMLString({
					name: 'mergeCells',
					attributes: [['count', this._mergedCells.length]],
					children: children
				});
			}
			return '';
		}
	}]);

	return MergedCells;
}(DrawingsExt);

module.exports = MergedCells;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":26,"./drawing":28}],32:[function(require,module,exports){
(function (global){
'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var SheetView = require('./sheetView');

var PrepareSave = function (_SheetView) {
	_inherits(PrepareSave, _SheetView);

	function PrepareSave() {
		_classCallCheck(this, PrepareSave);

		return _possibleConstructorReturn(this, (PrepareSave.__proto__ || Object.getPrototypeOf(PrepareSave)).apply(this, arguments));
	}

	_createClass(PrepareSave, [{
		key: '_prepare',
		value: function _prepare() {
			var maxX = 0;

			this.preparedData = [];
			this.preparedColumns = [];
			this.preparedRows = [];

			this._prepareTables();
			prepareColumns(this);
			prepareRows(this);

			for (var rowIndex = 0, len = this.data.length; rowIndex < len; rowIndex++) {
				var preparedDataRow = prepareDataRow(this, rowIndex);

				if (preparedDataRow.length > maxX) {
					maxX = preparedDataRow.length;
				}
			}

			this.data = null;
			this.columns = null;
			this.rows = null;

			this.maxX = maxX;
			this.maxY = this.preparedData.length;
		}
	}]);

	return PrepareSave;
}(SheetView);

function prepareColumns(worksheet) {
	var styles = worksheet.common.styles;

	_.forEach(worksheet.columns, function (column, index) {
		if (column) {
			var preparedColumn = _.clone(column);

			if (column.style) {
				var style = styles.addFormat(column.style);
				var columnStyle = styles._get(style).fillOut ? style : styles._addInvisibleFormat(style);

				preparedColumn.style = style;
				preparedColumn.styleId = styles._getId(columnStyle);
			}
			worksheet.preparedColumns[index] = preparedColumn;
		}
	});
}

function prepareRows(worksheet) {
	var styles = worksheet.common.styles;

	_.forEach(worksheet.rows, function (row, index) {
		if (row) {
			var preparedRow = _.clone(row);

			if (row.style) {
				preparedRow.style = styles.addFormat(row.style);
			}
			worksheet.preparedRows[index] = preparedRow;
		}
	});
}

function prepareDataRow(worksheet, rowIndex) {
	var styles = worksheet.common.styles;
	var strings = worksheet.common.strings;
	var preparedDataRow = [];
	var row = worksheet.preparedRows[rowIndex];
	var dataRow = worksheet.data[rowIndex];

	if (dataRow) {
		var rowStyle = null;

		if (!_.isArray(dataRow)) {
			row = mergeDataRowToRow(styles, row, dataRow);
			dataRow = dataRow.data;
		}
		if (row) {
			rowStyle = row.style || null;
		}

		for (var colIndex = 0; colIndex < dataRow.length; colIndex++) {
			var column = worksheet.preparedColumns[colIndex];
			var value = dataRow[colIndex];
			var cellValue = void 0;
			var cellType = void 0;
			var cellStyle = null;
			var cellFormula = null;
			var isString = false;

			if (_.isDate(value)) {
				cellValue = value;
				cellType = 'date';
			} else if (value && (typeof value === 'undefined' ? 'undefined' : _typeof(value)) === 'object') {
				if (value.style) {
					cellStyle = value.style;
				}
				if (value.formula) {
					cellValue = value.formula;
					cellType = 'formula';
				} else if (value.date) {
					cellValue = value.date;
					cellType = 'date';
				} else if (value.time) {
					cellValue = value.time;
					cellType = 'time';
				} else {
					cellValue = value.value;
					cellType = value.type;
				}

				insertEmbedded(worksheet, dataRow, value, colIndex, rowIndex);
			} else {
				cellValue = value;
				cellType = null;
			}

			cellStyle = styles._merge(column ? column.style : null, rowStyle, cellStyle);

			if (!cellType) {
				if (row && row.type) {
					cellType = row.type;
				} else if (column && column.type) {
					cellType = column.type;
				} else if (typeof cellValue === 'number') {
					cellType = 'number';
				} else if (typeof cellValue === 'string') {
					cellType = 'string';
				}
			}

			if (cellType === 'string') {
				cellValue = strings.add(cellValue);
				isString = true;
			} else if (cellType === 'date' || cellType === 'time') {
				var dateValue = _.isDate(cellValue) ? cellValue.valueOf() : cellValue;
				var date = 25569.0 + (dateValue - worksheet.timezoneOffset) / (60 * 60 * 24 * 1000);

				if (_.isFinite(date)) {
					cellValue = date;
				} else {
					cellValue = strings.add(String(cellValue));
					isString = true;
				}
			} else if (cellType === 'formula') {
				cellFormula = _.escape(cellValue);
				cellValue = null;
			}

			preparedDataRow[colIndex] = {
				value: cellValue,
				formula: cellFormula,
				styleId: styles._getId(cellStyle),
				isString: isString
			};
		}
	}

	worksheet.preparedData[rowIndex] = preparedDataRow;
	if (row) {
		setRowStyleId(styles, row);
		worksheet.preparedRows[rowIndex] = row;
	}

	return preparedDataRow;
}

function mergeDataRowToRow(styles) {
	var row = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};
	var dataRow = arguments[2];

	row.height = dataRow.height || row.height;
	row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;
	row.type = dataRow.type || row.type;
	row.style = dataRow.style ? styles.addFormat(dataRow.style) : row.style;

	return row;
}

function setRowStyleId(styles, row) {
	if (row.style) {
		var rowStyle = styles._get(row.style).fillOut ? row.style : styles._addInvisibleFormat(row.style);

		row.styleId = styles._getId(rowStyle);
	}
}

function insertEmbedded(worksheet, dataRow, value, colIndex, rowIndex) {
	if (value.hyperlink) {
		worksheet._insertHyperlink(colIndex, rowIndex, value.hyperlink);
	}

	if (value.image) {
		worksheet._insertDrawing(colIndex, rowIndex, value.image);
	}

	if (value.colspan || value.rowspan) {
		var colSpan = (value.colspan || 1) - 1;
		var rowSpan = (value.rowspan || 1) - 1;

		if (colSpan || rowSpan) {
			worksheet.mergeCells({ c: colIndex + 1, r: rowIndex + 1 }, { c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan });
			worksheet._insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan);
		}
	}
}

module.exports = PrepareSave;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./sheetView":35}],33:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

var Print = function () {
	function Print() {
		_classCallCheck(this, Print);

		this._headers = [];
		this._footers = [];
	}
	/**
  * Expects an array length of three.
  * @param {Array} headers [left, center, right]
  */


	_createClass(Print, [{
		key: 'setHeader',
		value: function setHeader(headers) {
			if (!_.isArray(headers)) {
				throw 'Invalid argument type - setHeader expects an array of three instructions';
			}
			this._headers = headers;
			return this;
		}
		/**
   * Expects an array length of three.
   * @param {Array} footers [left, center, right]
   */

	}, {
		key: 'setFooter',
		value: function setFooter(footers) {
			if (!_.isArray(footers)) {
				throw 'Invalid argument type - setFooter expects an array of three instructions';
			}
			this._footers = footers;
			return this;
		}
		/**
   * Set page details in inches.
   */

	}, {
		key: 'setPageMargin',
		value: function setPageMargin(margin) {
			this._margin = _.defaults(margin, {
				left: 0.7,
				right: 0.7,
				top: 0.75,
				bottom: 0.75,
				header: 0.3,
				footer: 0.3
			});
			return this;
		}
		/**
   * http://www.datypic.com/sc/ooxml/t-ssml_ST_Orientation.html
   *
   * Can be one of 'portrait' or 'landscape'.
   *
   * @param {String} orientation
   */

	}, {
		key: 'setPageOrientation',
		value: function setPageOrientation(orientation) {
			this._orientation = orientation;
			return this;
		}
		/**
   * Set rows to repeat for print
   *
   * @param {int|[int, int]} params - number of rows to repeat from the top | [first, last] repeat rows
   */

	}, {
		key: 'setPrintTitleTop',
		value: function setPrintTitleTop(params) {
			this._printTitles = this._printTitles || {};

			if (_.isObject(params)) {
				this._printTitles.topFrom = params[0];
				this._printTitles.topTo = params[1];
			} else {
				this._printTitles.topFrom = 0;
				this._printTitles.topTo = params - 1;
			}
			return this;
		}
		/**
   * Set columns to repeat for print
   *
   * @param {int|[int, int]} params - number of columns to repeat from the left | [first, last] repeat columns
   */

	}, {
		key: 'setPrintTitleLeft',
		value: function setPrintTitleLeft(params) {
			this._printTitles = this._printTitles || {};

			if (_.isObject(params)) {
				this._printTitles.leftFrom = params[0];
				this._printTitles.leftTo = params[1];
			} else {
				this._printTitles.leftFrom = 0;
				this._printTitles.leftTo = params - 1;
			}
			return this;
		}
	}, {
		key: '_savePrint',
		value: function _savePrint() {
			return savePageMargins(this._margin) + savePageSetup(this._orientation) + saveHeaderFooter(this._headers, this._footers);
		}
	}]);

	return Print;
}();

function savePageMargins(margin) {
	if (margin) {
		return toXMLString({
			name: 'pageMargins',
			attributes: [['top', margin.top], ['bottom', margin.bottom], ['left', margin.left], ['right', margin.right], ['header', margin.header], ['footer', margin.footer]]
		});
	}
	return '';
}

function savePageSetup(orientation) {
	if (orientation) {
		return toXMLString({
			name: 'pageSetup',
			attributes: [['orientation', orientation]]
		});
	}
	return '';
}

function saveHeaderFooter(headers, footers) {
	if (headers.length > 0 || footers.length > 0) {
		var children = [];

		if (headers.length > 0) {
			children.push(saveHeader(headers));
		}
		if (footers.length > 0) {
			children.push(saveFooter(footers));
		}

		return toXMLString({
			name: 'headerFooter',
			children: children
		});
	}
	return '';
}

function saveHeader(headers) {
	return toXMLString({
		name: 'oddHeader',
		value: compilePageDetailPackage(headers)
	});
}

function saveFooter(footers) {
	return toXMLString({
		name: 'oddFooter',
		value: compilePageDetailPackage(footers)
	});
}

function compilePageDetailPackage(data) {
	data = data || '';

	return ['&L', compilePageDetailPiece(data[0] || ''), '&C', compilePageDetailPiece(data[1] || ''), '&R', compilePageDetailPiece(data[2] || '')].join('');
}

function compilePageDetailPiece(data) {
	if (_.isString(data)) {
		return '&"-,Regular"'.concat(data);
	} else if (_.isObject(data) && !_.isArray(data)) {
		var string = '';

		if (data.font || data.bold) {
			var weighting = data.bold ? 'Bold' : 'Regular';

			string += '&"' + (data.font || '-') + ',' + weighting + '"';
		} else {
			string += '&"-,Regular"';
		}
		if (data.underline) {
			string += '&U';
		}
		if (data.fontSize) {
			string += '&' + data.fontSize;
		}
		string += data.text;

		return string;
	} else if (_.isArray(data)) {
		return _.reduce(data, function (result, value) {
			return result.concat(compilePageDetailPiece(value));
		}, '');
	}
}

module.exports = Print;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],34:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var Readable = require('stream').Readable;
var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var PrepareSave = require('./prepareSave');
var util = require('../util');
var toXMLString = require('../XMLString');

var WorksheetSave = function (_PrepareSave) {
	_inherits(WorksheetSave, _PrepareSave);

	function WorksheetSave() {
		_classCallCheck(this, WorksheetSave);

		return _possibleConstructorReturn(this, (WorksheetSave.__proto__ || Object.getPrototypeOf(WorksheetSave)).apply(this, arguments));
	}

	_createClass(WorksheetSave, [{
		key: '_save',
		value: function _save(canStream) {
			if (canStream) {
				return new WorksheetStream({
					worksheet: this
				});
			} else {
				return saveBeforeRows(this) + saveData(this) + saveAfterRows(this);
			}
		}
	}]);

	return WorksheetSave;
}(PrepareSave);

var WorksheetStream = function (_ref) {
	_inherits(WorksheetStream, _ref);

	function WorksheetStream(options) {
		_classCallCheck(this, WorksheetStream);

		var _this2 = _possibleConstructorReturn(this, (WorksheetStream.__proto__ || Object.getPrototypeOf(WorksheetStream)).call(this, options));

		_this2.worksheet = options.worksheet;
		_this2.status = 0;
		return _this2;
	}

	_createClass(WorksheetStream, [{
		key: '_read',
		value: function _read(size) {
			var worksheet = this.worksheet;
			var stop = false;

			if (this.status === 0) {
				stop = !this.push(saveBeforeRows(worksheet));

				this.status = 1;
				this.index = 0;
				this.len = worksheet.preparedData.length;
			}

			if (this.status === 1) {
				var data = worksheet.preparedData;
				var preparedRows = worksheet.preparedRows;
				var s = '';

				while (this.index < this.len && !stop) {
					while (this.index < this.len && s.length < size) {
						s += saveRow(data[this.index], preparedRows[this.index], this.index);
						data[this.index] = null;
						this.index++;
					}
					stop = !this.push(s);
					s = '';
				}

				if (this.index === this.len) {
					this.status = 2;
				}
			}

			if (this.status === 2) {
				this.push(saveAfterRows(worksheet));
				this.push(null);
			}
		}
	}]);

	return WorksheetStream;
}(Readable || null);

function saveBeforeRows(worksheet) {
	return util.xmlPrefix + '<worksheet xmlns="' + util.schemas.spreadsheetml + '" xmlns:r="' + util.schemas.relationships + '" xmlns:mc="' + util.schemas.markupCompat + '">' + saveDimension(worksheet.maxX, worksheet.maxY) + worksheet._saveSheetView() + saveColumns(worksheet.preparedColumns) + '<sheetData>';
}

function saveAfterRows(worksheet) {
	return '</sheetData>' +
	// 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
	// with Microsoft Excel (2007, 2013)
	worksheet._saveMergeCells() + worksheet._saveHyperlinks() + worksheet._savePrint() + worksheet._saveTables() +
	// the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
	// to issue with Microsoft Excel (2007, 2013)
	worksheet._saveDrawing() + '</worksheet>';
}

function saveData(worksheet) {
	var data = worksheet.preparedData;
	var preparedRows = worksheet.preparedRows;
	var children = '';

	for (var i = 0, len = data.length; i < len; i++) {
		children += saveRow(data[i], preparedRows[i], i);
		data[i] = null;
	}
	return children;
}

function saveRow(dataRow, row, rowIndex) {
	var rowChildren = [];

	if (dataRow) {
		var rowLen = dataRow.length;
		rowChildren = new Array(rowLen);

		for (var colIndex = 0; colIndex < rowLen; colIndex++) {
			var value = dataRow[colIndex];

			if (!value) {
				continue;
			}

			var attrs = ' r="' + util.positionToLetter(colIndex + 1, rowIndex + 1) + '"';
			if (value.styleId) {
				attrs += ' s="' + value.styleId + '"';
			}
			if (value.isString) {
				attrs += ' t="s"';
			}

			if (!_.isNil(value.value)) {
				rowChildren[colIndex] = '<c' + attrs + '><v>' + value.value + '</v></c>';
			} else if (!_.isNil(value.formula)) {
				rowChildren[colIndex] = '<c' + attrs + '><f>' + value.formula + '</f></c>';
			} else {
				rowChildren[colIndex] = '<c' + attrs + '/>';
			}
		}
	}

	return '<row' + getRowAttributes(row, rowIndex) + '>' + rowChildren.join('') + '</row>';
}

function getRowAttributes(row, rowIndex) {
	var attributes = ' r="' + (rowIndex + 1) + '"';

	if (row) {
		if (row.height !== undefined) {
			attributes += ' customHeight="1"';
			attributes += ' ht="' + row.height + '"';
		}
		if (row.styleId) {
			attributes += ' customFormat="1"';
			attributes += ' s="' + row.styleId + '"';
		}
		if (row.outlineLevel) {
			attributes += ' outlineLevel="' + row.outlineLevel + '"';
		}
	}
	return attributes;
}

function saveDimension(maxX, maxY) {
	var attributes = [];

	if (maxX !== 0) {
		attributes.push(['ref', 'A1:' + util.positionToLetter(maxX, maxY)]);
	} else {
		attributes.push(['ref', 'A1']);
	}

	return toXMLString({
		name: 'dimension',
		attributes: attributes
	});
}

function saveColumns(columns) {
	if (columns.length) {
		var children = _.map(columns, function (column, index) {
			column = column || {};

			var attributes = [['min', column.min || index + 1], ['max', column.max || index + 1]];

			if (column.hidden) {
				attributes.push(['hidden', 1]);
			}
			if (column.bestFit) {
				attributes.push(['bestFit', 1]);
			}
			if (column.customWidth || column.width) {
				attributes.push(['customWidth', 1]);
			}
			if (column.width) {
				attributes.push(['width', column.width]);
			} else {
				attributes.push(['width', 9.140625]);
			}
			if (column.styleId) {
				attributes.push(['style', column.styleId]);
			}

			return toXMLString({
				name: 'col',
				attributes: attributes
			});
		});

		return toXMLString({
			name: 'cols',
			children: children
		});
	}
	return '';
}

module.exports = WorksheetSave;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":26,"./prepareSave":32,"stream":1}],35:[function(require,module,exports){
(function (global){
'use strict';

/**
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sheetview.aspx
 */

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var Hyperlinks = require('./hyperlinks');
var toXMLString = require('../XMLString');

var SheetView = function (_Hyperlinks) {
	_inherits(SheetView, _Hyperlinks);

	function SheetView(workbook, config) {
		_classCallCheck(this, SheetView);

		var _this = _possibleConstructorReturn(this, (SheetView.__proto__ || Object.getPrototypeOf(SheetView)).call(this, workbook, config));

		_this._pane = null;
		_this._attributes = {
			defaultGridColor: {
				value: null,
				bool: true
			},
			colorId: {
				value: null
			},
			rightToLeft: {
				value: null,
				bool: true
			},
			showFormulas: {
				value: null,
				bool: true
			},
			showGridLines: {
				value: null,
				bool: true
			},
			showOutlineSymbols: {
				value: null,
				bool: true
			},
			showRowColHeaders: {
				value: null,
				bool: true
			},
			showRuler: {
				value: null,
				bool: true
			},
			showWhiteSpace: {
				value: null,
				bool: true
			},
			showZeros: {
				value: null,
				bool: true
			},
			tabSelected: {
				value: null,
				bool: true
			},
			topLeftCell: {
				value: null //A1
			},
			view: {
				value: 'normal' //normal | pageBreakPreview | pageLayout
			},
			windowProtection: {
				value: null,
				bool: true
			},
			zoomScale: {
				value: null //10-400
			},
			zoomScaleNormal: {
				value: null //10-400
			},
			zoomScalePageLayoutView: {
				value: null //10-400
			},
			zoomScaleSheetLayoutView: {
				value: null //10-400
			}
		};

		_.forEach(config, function (value, name) {
			if (name === 'freeze') {
				_this.freeze(value.col, value.row, value.cell, value.activePane);
			} else if (name === 'split') {
				_this.split(value.x, value.y, value.cell, value.activePane);
			} else if (_this._attributes[name]) {
				_this._attributes[name].value = value;
			}
		});
		return _this;
	}

	_createClass(SheetView, [{
		key: 'setAttribute',
		value: function setAttribute(name, value) {
			if (this._attributes[name]) {
				this._attributes[name].value = value;
			}
			return this;
		}
		/**
   * Add froze pane
   * @param col - column number: 0, 1, 2 ...
   * @param row - row number: 0, 1, 2 ...
   * @param cell? - 'A1' | {c: 1, r: 1}
   * @param activePane? - topLeft | topRight | bottomLeft | bottomRight
   */

	}, {
		key: 'freeze',
		value: function freeze(col, row, cell, activePane) {
			this._pane = {
				state: 'frozen',
				xSplit: col,
				ySplit: row,
				topLeftCell: util.canonCell(cell) || util.positionToLetter(col + 1, row + 1),
				activePane: activePane || 'bottomRight'
			};
			return this;
		}
		/**
   * Add split pane
   * @param x - Horizontal position of the split, in points; 0 (zero) if none
   * @param y - Vertical position of the split, in points; 0 (zero) if none
   * @param cell? - 'A1' | {c: 1, r: 1}
   * @param activePane? - topLeft | topRight | bottomLeft | bottomRight
   */

	}, {
		key: 'split',
		value: function split(x, y, cell, activePane) {
			this._pane = {
				state: 'split',
				xSplit: x * 20,
				ySplit: y * 20,
				topLeftCell: util.canonCell(cell) || 'A1',
				activePane: activePane || 'bottomRight'
			};
			return this;
		}
	}, {
		key: '_saveSheetView',
		value: function _saveSheetView() {
			var attributes = [['workbookViewId', 0]];

			_.forEach(this._attributes, function (attr, name) {
				var value = attr.value;

				if (value !== null) {
					if (attr.bool) {
						value = value ? 'true' : 'false';
					}
					attributes.push([name, value]);
				}
			});

			return toXMLString({
				name: 'sheetViews',
				children: [toXMLString({
					name: 'sheetView',
					attributes: attributes,
					children: [savePane(this._pane)]
				})]
			});
		}
	}]);

	return SheetView;
}(Hyperlinks);

function savePane(pane) {
	if (pane) {
		return toXMLString({
			name: 'pane',
			attributes: [['state', pane.state], ['xSplit', pane.xSplit], ['ySplit', pane.ySplit], ['topLeftCell', pane.topLeftCell], ['activePane', pane.activePane]]
		});
	}
	return '';
}

module.exports = SheetView;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":26,"./hyperlinks":29}],36:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var Print = require('./print');
var Table = require('../table');
var toXMLString = require('../XMLString');

var Tables = function (_Print) {
	_inherits(Tables, _Print);

	function Tables() {
		_classCallCheck(this, Tables);

		var _this = _possibleConstructorReturn(this, (Tables.__proto__ || Object.getPrototypeOf(Tables)).call(this));

		_this._tables = [];
		return _this;
	}

	_createClass(Tables, [{
		key: 'addTable',
		value: function addTable(config) {
			var table = new Table(this, config);

			this.common.addTable(table);
			this.relations.addRelation(table, 'table');
			this._tables.push(table);

			return table;
		}
	}, {
		key: '_prepareTables',
		value: function _prepareTables() {
			var data = this.data;

			_.forEach(this._tables, function (table) {
				table._prepare(data);
			});
		}
	}, {
		key: '_saveTables',
		value: function _saveTables() {
			var _this2 = this;

			if (this._tables.length > 0) {
				var children = _.map(this._tables, function (table) {
					return toXMLString({
						name: 'tablePart',
						attributes: [['r:id', _this2.relations.getRelationshipId(table)]]
					});
				});

				return toXMLString({
					name: 'tableParts',
					attributes: [['count', this._tables.length]],
					children: children
				});
			}
			return '';
		}
	}]);

	return Tables;
}(Print);

module.exports = Tables;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../table":25,"./print":33}]},{},[11])(11)
});