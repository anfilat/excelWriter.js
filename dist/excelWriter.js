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
},{"./util":29}],3:[function(require,module,exports){
'use strict';

function Images(common) {
	this.common = common;
	this.images = Object.create(null);
	this.imageByNames = Object.create(null);
	this.extensions = Object.create(null);
}

Images.prototype = {
	add: function add(data) {
		var type = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : '';
		var name = arguments[2];

		var image = this.images[data];

		if (!image) {
			var id = this.common.uniqueIdForSpace('image');
			var path = 'xl/media/image' + id + '.' + type;
			var contentType = getContentType(type);

			name = name || 'excelWriter' + id;
			image = {
				objectId: 'image' + id,
				data: data,
				name: name,
				contentType: contentType,
				extension: type,
				path: path
			};
			this.common.addPath(image, '/' + path);
			this.images[data] = image;
			this.imageByNames[name] = image;
			this.extensions[type] = contentType;
		} else if (name && !this.imageByNames[name]) {
			image.name = name;
			this.imageByNames[name] = image;
		}
		return image.name;
	},
	get: function get(name) {
		return this.imageByNames[name];
	},
	getImages: function getImages() {
		return this.images;
	},
	removeImages: function removeImages() {
		this.images = null;
		this.imageByNames = null;
	},
	getExtensions: function getExtensions() {
		return this.extensions;
	}
};

function getContentType(type) {
	switch (type.toLowerCase()) {
		case 'jpeg':
		case 'jpg':
			return 'image/jpeg';
		case 'png':
			return 'image/png';
		case 'gif':
			return 'image/gif';
		default:
			return null;
	}
}

module.exports = Images;

},{}],4:[function(require,module,exports){
'use strict';

var Images = require('./images');
var SharedStrings = require('./sharedStrings');
var Styles = require('../styles');

function Common() {
	this.idSpaces = Object.create(null);
	this.paths = Object.create(null);

	this.images = new Images(this);

	this.strings = new SharedStrings(this);
	this.addPath(this.strings, 'sharedStrings.xml');

	this.styles = new Styles(this);
	this.addPath(this.styles, 'styles.xml');

	this.worksheets = [];
	this.tables = [];
	this.drawings = [];
}

Common.prototype = {
	uniqueId: function uniqueId(space) {
		return space + this.uniqueIdForSpace(space);
	},
	uniqueIdForSpace: function uniqueIdForSpace(space) {
		if (!this.idSpaces[space]) {
			this.idSpaces[space] = 1;
		}
		return this.idSpaces[space]++;
	},
	addPath: function addPath(object, path) {
		this.paths[object.objectId] = path;
	},
	getPath: function getPath(object) {
		return this.paths[object.objectId];
	},
	addWorksheet: function addWorksheet(worksheet) {
		var index = this.worksheets.length + 1;
		var path = 'worksheets/sheet' + index + '.xml';
		var relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

		worksheet.path = 'xl/' + path;
		worksheet.relationsPath = relationsPath;
		this.worksheets.push(worksheet);
		this.addPath(worksheet, path);
	},
	getNewWorksheetDefaultName: function getNewWorksheetDefaultName() {
		return 'Sheet ' + (this.worksheets.length + 1);
	},
	setActiveWorksheet: function setActiveWorksheet(worksheet) {
		this.activeWorksheet = worksheet;
	},
	getActiveWorksheetIndex: function getActiveWorksheetIndex() {
		if (this.activeWorksheet) {
			var activeWorksheetId = this.activeWorksheet.objectId;

			return Math.max(0, this.worksheets.findIndex(function (worksheet) {
				return worksheet.objectId === activeWorksheetId;
			}));
		}
		return 0;
	},
	addTable: function addTable(table) {
		var index = this.tables.length + 1;
		var path = 'xl/tables/table' + index + '.xml';

		table.path = path;
		this.tables.push(table);
		this.addPath(table, '/' + path);
	},
	addDrawings: function addDrawings(drawings) {
		var index = this.drawings.length + 1;
		var path = 'xl/drawings/drawing' + index + '.xml';
		var relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

		drawings.path = path;
		drawings.relationsPath = relationsPath;
		this.drawings.push(drawings);
		this.addPath(drawings, '/' + path);
	}
};

module.exports = Common;

},{"../styles":19,"./images":3,"./sharedStrings":5}],5:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var Readable = require('stream').Readable;
var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');

var spaceRE = /^\s|\s$/;

function SharedStrings(common) {
	this.objectId = common.uniqueId('SharedStrings');
	this.strings = Object.create(null);
	this.stringArray = [];
	this.count = 0;
}

SharedStrings.prototype = {
	add: function add(string) {
		var stringId = this.strings[string];

		if (stringId === undefined) {
			stringId = this.count++;

			this.strings[string] = stringId;
			this.stringArray[stringId] = string;
		}

		return stringId;
	},
	isStrings: function isStrings() {
		return this.count > 0;
	},
	save: function save() {
		var canStream = !!Readable;

		this.strings = null;

		if (canStream) {
			return new SharedStringsStream({
				strings: this.stringArray
			});
		}
		var result = getXMLBegin(this.count) + this.stringArray.map(function (string) {
			string = _.escape(string);

			if (spaceRE.test(string)) {
				return '<si><t xml:space="preserve">' + string + '</t></si>';
			}
			return '<si><t>' + string + '</t></si>';
		}).join('') + getXMLEnd();
		this.stringArray = null;
		return result;
	}
};

var SharedStringsStream = function (_ref) {
	_inherits(SharedStringsStream, _ref);

	function SharedStringsStream(options) {
		_classCallCheck(this, SharedStringsStream);

		var _this = _possibleConstructorReturn(this, (SharedStringsStream.__proto__ || Object.getPrototypeOf(SharedStringsStream)).call(this, options));

		_this.strings = options.strings;
		_this.status = 0;
		_this.index = 0;
		_this.len = _this.strings.length;
		return _this;
	}

	_createClass(SharedStringsStream, [{
		key: '_read',
		value: function _read(size) {
			var stop = false;

			if (this.status === 0) {
				stop = !this.push(getXMLBegin(this.len));

				this.status = 1;
			}

			if (this.status === 1) {
				while (this.index < this.len && !stop) {
					stop = !this.push(this.packChunk(size));
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
	}, {
		key: 'packChunk',
		value: function packChunk(size) {
			var s = '';

			while (this.index < this.len && s.length < size) {
				var str = _.escape(this.strings[this.index]);

				if (spaceRE.test(str)) {
					s += '<si><t xml:space="preserve">' + str + '</t></si>';
				} else {
					s += '<si><t>' + str + '</t></si>';
				}
				this.strings[this.index] = null;
				this.index++;
			}
			return s;
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
},{"../util":29,"stream":1}],6:[function(require,module,exports){
'use strict';

var borderStyle = {
	dashDot: 'dashDot',
	dashDotDot: 'dashDotDot',
	dashed: 'dashed',
	dotted: 'dotted',
	double: 'double',
	hair: 'hair',
	medium: 'medium',
	mediumDashDot: 'mediumDashDot',
	mediumDashDotDot: 'mediumDashDotDot',
	mediumDashed: 'mediumDashed',
	none: 'none',
	slantDashDot: 'slantDashDot',
	thick: 'thick',
	thin: 'thin'
};

module.exports = { borderStyle: borderStyle };

},{}],7:[function(require,module,exports){
(function (global){
'use strict';

// https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.spreadsheet.twocellanchor.aspx

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

function Anchor(config) {
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

Anchor.prototype = {
	saveWithContent: function saveWithContent(content) {
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
};

module.exports = Anchor;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":29}],8:[function(require,module,exports){
'use strict';

// https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.spreadsheet.absoluteanchor.aspx

var util = require('../util');
var toXMLString = require('../XMLString');

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

	this.x = util.pixelsToEMUs(left);
	this.y = util.pixelsToEMUs(top);
	this.width = util.pixelsToEMUs(width);
	this.height = util.pixelsToEMUs(height);
}

AnchorAbsolute.prototype = {
	saveWithContent: function saveWithContent(content) {
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
};

module.exports = AnchorAbsolute;

},{"../XMLString":2,"../util":29}],9:[function(require,module,exports){
(function (global){
'use strict';

// https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.spreadsheet.onecellanchor.aspx

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

function AnchorOneCell(config) {
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

AnchorOneCell.prototype = {
	saveWithContent: function saveWithContent(content) {
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
};

module.exports = AnchorOneCell;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":29}],10:[function(require,module,exports){
'use strict';

var Relations = require('../relations');
var util = require('../util');
var toXMLString = require('../XMLString');
var Picture = require('./picture');

function Drawings(common) {
	this.common = common;

	this.objectId = this.common.uniqueId('Drawings');
	this.drawings = [];
	this.relations = new Relations(common);
}

Drawings.prototype = {
	add: function add(name, config, anchorType) {
		var image = this.common.images.get(name);
		var imageRelationId = this.relations.add(image, 'image');
		var picture = new Picture(this.common, {
			image: image,
			imageRelationId: imageRelationId,
			config: config,
			anchorType: anchorType
		});

		this.drawings.push(picture);
	},
	save: function save() {
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
};

module.exports = Drawings;

},{"../XMLString":2,"../relations":13,"../util":29,"./picture":11}],11:[function(require,module,exports){
'use strict';

// https://msdn.microsoft.com/en-us/library/documentformat.openxml.drawing.spreadsheet.picture.aspx

var util = require('../util');
var toXMLString = require('../XMLString');
var Anchor = require('./anchor');
var AnchorOneCell = require('./anchorOneCell');
var AnchorAbsolute = require('./anchorAbsolute');

function Picture(common, config) {
	this.pictureId = common.uniqueIdForSpace('Picture');
	this.image = config.image;
	this.imageRelationId = config.imageRelationId;
	this.createAnchor(config.anchorType, config.config);
}

Picture.prototype = {
	/**
  *
  * @param {String} type Can be 'anchor', 'oneCell' or 'absolute'.
  * @param {Object} config Shorthand - pass the created anchor coords that can normally be used to construct it.
  */
	createAnchor: function createAnchor(type, config) {
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
	},
	save: function save() {
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
};

module.exports = Picture;

},{"../XMLString":2,"../util":29,"./anchor":7,"./anchorAbsolute":8,"./anchorOneCell":9}],12:[function(require,module,exports){
'use strict';

var Workbook = require('./workbook');
var constants = require('./constants');

// Excel workbook API
function createWorkbook() {
	var outerWorkbook = {
		addWorksheet: function addWorksheet(config) {
			return workbook.addWorksheet(config);
		},
		addFormat: function addFormat(format, name) {
			return styles.addFormat(format, name);
		},
		addFontFormat: function addFontFormat(format, name) {
			return styles.addFontFormat(format, name);
		},
		addBorderFormat: function addBorderFormat(format, name) {
			return styles.addBorderFormat(format, name);
		},
		addPatternFormat: function addPatternFormat(format, name) {
			return styles.addPatternFormat(format, name);
		},
		addGradientFormat: function addGradientFormat(format, name) {
			return styles.addGradientFormat(format, name);
		},
		addNumberFormat: function addNumberFormat(format, name) {
			return styles.addNumberFormat(format, name);
		},
		addTableFormat: function addTableFormat(format, name) {
			return styles.addTableFormat(format, name);
		},
		addTableElementFormat: function addTableElementFormat(format, name) {
			return styles.addTableElementFormat(format, name);
		},
		setDefaultTableStyle: function setDefaultTableStyle(name) {
			styles.setDefaultTableStyle(name);
			return this;
		},
		addImage: function addImage(data, type, name) {
			return images.add(data, type, name);
		},
		save: function save(options) {
			return workbook.save(options);
		},
		saveAsNodeStream: function saveAsNodeStream(options) {
			return workbook.saveAsNodeStream(options);
		}
	};
	var workbook = new Workbook(outerWorkbook);
	var styles = workbook.styles;
	var images = workbook.images;

	return outerWorkbook;
}

module.exports = constants;
module.exports.createWorkbook = createWorkbook;

},{"./constants":6,"./workbook":30}],13:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');
var toXMLString = require('./XMLString');

function RelationshipManager(common) {
	this.common = common;

	this.relations = Object.create(null);
	this.lastId = 1;
}

RelationshipManager.prototype = {
	add: function add(object, type) {
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
	},
	getId: function getId(object) {
		var relation = this.relations[object.objectId];

		return relation ? relation.relationId : null;
	},
	save: function save() {
		var _this = this;

		var children = _.map(this.relations, function (relation) {
			var attributes = [['Id', relation.relationId], ['Type', relation.schema], ['Target', relation.object.target || _this.common.getPath(relation.object)]];

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
};

module.exports = RelationshipManager;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./util":29}],14:[function(require,module,exports){
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
},{"../XMLString":2}],15:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;

var _require = require('../constants'),
    borderStyle = _require.borderStyle;

var StylePart = require('./stylePart');

var _require2 = require('./utils'),
    saveColor = _require2.saveColor;

var toXMLString = require('../XMLString');

var MAIN_BORDERS = ['left', 'right', 'top', 'bottom'];
var BORDERS = ['left', 'right', 'top', 'bottom', 'diagonal'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borders.aspx
function Borders(styles) {
	StylePart.call(this, styles, 'borders', 'border');

	this.init();
	this.lastId = this.formats.length;
}

Borders.canon = function (format) {
	var result = {};

	if (_.has(format, 'style') || _.has(format, 'color')) {
		if (!format.style || borderStyle[format.style]) {
			MAIN_BORDERS.forEach(function (name) {
				result[name] = {
					style: format.style,
					color: format.color
				};
			});
		}
	} else {
		BORDERS.forEach(function (name) {
			var border = format[name];

			if (border && (!format.style || borderStyle[border.style])) {
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
};

Borders.saveFormat = function (format) {
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
};

Borders.prototype = _.merge({}, StylePart.prototype, {
	init: function init() {
		this.formats.push({ format: this.canon({}) });
	},

	canon: Borders.canon,
	saveFormat: Borders.saveFormat,
	merge: function merge() {
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
});

module.exports = Borders;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../constants":6,"./stylePart":23,"./utils":26}],16:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var alignment = require('./alignment');
var protection = require('./protection');
var PREDEFINED = require('./predeinedFormats');
var toXMLString = require('../XMLString');

var ALLOWED_PARTS = ['format', 'fill', 'border', 'font'];
var XLS_NAMES = ['numFmtId', 'fillId', 'borderId', 'fontId'];

function Cells(styles) {
	StylePart.call(this, styles, 'cellXfs', 'format');

	this.init();
	this.lastId = this.formats.length;
	this.saveEmpty = false;
}

Cells.prototype = _.merge({}, StylePart.prototype, {
	init: function init() {
		var _this = this;

		this.formats.push({ format: this.canon({}) });

		_.keys(PREDEFINED).forEach(function (name) {
			_this.predefined[name] = { format: name };
		});
	},
	canon: function canon(format, flags) {
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
		var alignmentValue = flags && flags.merge ? alignment.canon(format.alignment) : alignment.canon(format);
		if (alignmentValue) {
			result.alignment = alignmentValue;
		}
		var protectionValue = flags && flags.merge ? protection.canon(format.protection) : protection.canon(format);
		if (protectionValue) {
			result.protection = protectionValue;
		}
		if (format.fillOut) {
			result.fillOut = format.fillOut;
		}
		return result;
	},
	fullGet: function fullGet(format) {
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
	},
	cutVisible: function cutVisible(format) {
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
	},
	merge: function merge(formatTo, formatFrom) {
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
	},
	saveFormat: function saveFormat(format) {
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
});

module.exports = Cells;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./alignment":14,"./predeinedFormats":21,"./protection":22,"./stylePart":23}],17:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');

var _require = require('../util'),
    canonColor = _require.canonColor;

var _require2 = require('./utils'),
    saveColor = _require2.saveColor;

var toXMLString = require('../XMLString');

var PATTERN_TYPES = ['none', 'solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fills.aspx
function Fills(styles) {
	StylePart.call(this, styles, 'fills', 'fill');

	this.init();
	this.lastId = this.formats.length;
}

Fills.canon = function (format, flags) {
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
};

Fills.saveFormat = function (format) {
	var children = format.fillType === 'pattern' ? [savePatternFill(format)] : [saveGradientFill(format)];

	return toXMLString({
		name: 'fill',
		children: children
	});
};

Fills.prototype = _.merge({}, StylePart.prototype, {
	init: function init() {
		this.formats.push({ format: this.canon({ type: 'none' }, { fillType: 'pattern' }) }, { format: this.canon({ type: 'gray125' }, { fillType: 'pattern' }) });
	},

	canon: Fills.canon,
	saveFormat: Fills.saveFormat,
	merge: function merge(formatTo, formatFrom) {
		return formatFrom || formatTo;
	}
});

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
},{"../XMLString":2,"../util":29,"./stylePart":23,"./utils":26}],18:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');

var _require = require('./utils'),
    saveColor = _require.saveColor;

var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fonts.aspx
function Fonts(styles) {
	StylePart.call(this, styles, 'fonts', 'font');

	this.init();
	this.lastId = this.formats.length;
}

Fonts.canon = function (format) {
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
};

Fonts.saveFormat = function (format) {
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
};

Fonts.prototype = _.merge({}, StylePart.prototype, {
	init: function init() {
		this.formats.push({ format: this.canon({}) });
	},

	canon: Fonts.canon,
	saveFormat: Fonts.saveFormat,
	merge: function merge(formatTo, formatFrom) {
		var result = _.assign(formatTo, formatFrom);

		result.color = formatFrom && formatFrom.color || formatTo && formatTo.color;
		return result;
	}
});

module.exports = Fonts;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./stylePart":23,"./utils":26}],19:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var NumberFormats = require('./numberFormats');
var Fonts = require('./fonts');
var Fills = require('./fills');
var Borders = require('./borders');
var Cells = require('./cells');
var Tables = require('./tables');
var TableElements = require('./tableElements');
var toXMLString = require('../XMLString');

function Styles(common) {
	this.objectId = common.uniqueId('Styles');
	this.numberFormats = new NumberFormats(this);
	this.fonts = new Fonts(this);
	this.fills = new Fills(this);
	this.borders = new Borders(this);
	this.cells = new Cells(this);
	this.tableElements = new TableElements(this);
	this.tables = new Tables(this);
	this.defaultTableStyle = '';
	this.mergeCache = Object.create(null);
}

Styles.prototype = {
	addFormat: function addFormat(format, name) {
		return this.cells.add(format, name);
	},
	get: function get(name) {
		return this.cells.get(name);
	},
	getId: function getId(name) {
		return this.cells.getId(name);
	},
	addFillOutFormat: function addFillOutFormat(format) {
		if (this.get(format).fillOut) {
			return format;
		}
		var style = this.cells.cutVisible(this.cells.fullGet(format));

		if (!_.isEmpty(style)) {
			return this.addFormat(style);
		}
	},
	merge: function merge() {
		var _this = this;

		for (var _len = arguments.length, formats = Array(_len), _key = 0; _key < _len; _key++) {
			formats[_key] = arguments[_key];
		}

		formats = _.compact(formats);

		if (formats.length === 0) {
			return null;
		} else if (formats.length === 1) {
			return this.addFormat(formats[0]);
		} else {
			var id = formats.reduce(function (result, format) {
				return result + JSON.stringify(format);
			}, '');
			var merged = this.mergeCache[id];

			if (!merged) {
				var newFormat = {};
				formats.forEach(function (format) {
					newFormat = _this.cells.merge(newFormat, _this.cells.fullGet(format));
				});
				merged = this.cells.add(newFormat, null, { merge: true });
				this.mergeCache[id] = merged;
			}
			return merged;
		}
	},
	addFontFormat: function addFontFormat(format, name) {
		return this.fonts.add(format, name);
	},
	addBorderFormat: function addBorderFormat(format, name) {
		return this.borders.add(format, name);
	},
	addPatternFormat: function addPatternFormat(format, name) {
		return this.fills.add(format, name, { fillType: 'pattern' });
	},
	addGradientFormat: function addGradientFormat(format, name) {
		return this.fills.add(format, name, { fillType: 'gradient' });
	},
	addNumberFormat: function addNumberFormat(format, name) {
		return this.numberFormats.add(format, name);
	},
	addTableFormat: function addTableFormat(format, name) {
		return this.tables.add(format, name);
	},
	addTableElementFormat: function addTableElementFormat(format, name) {
		return this.tableElements.add(format, name);
	},
	setDefaultTableStyle: function setDefaultTableStyle(name) {
		this.tables.defaultTableStyle = name;
	},
	save: function save() {
		return toXMLString({
			name: 'styleSheet',
			ns: 'spreadsheetml',
			children: [this.numberFormats.save(), this.fonts.save(), this.fills.save(), this.borders.save(), this.cells.save(), this.tableElements.save(), this.tables.save()]
		});
	}
};

module.exports = Styles;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./borders":15,"./cells":16,"./fills":17,"./fonts":18,"./numberFormats":20,"./tableElements":24,"./tables":25}],20:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var PREDEFINED = require('./predeinedFormats');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats.aspx
function NumberFormats(styles) {
	StylePart.call(this, styles, 'numFmts', 'numberFormat');

	this.init();
	this.lastId = 164;
}

NumberFormats.canon = function (format) {
	return format;
};

NumberFormats.saveFormat = function (format, styleFormat) {
	var attributes = [['numFmtId', styleFormat.formatId], ['formatCode', format]];

	return toXMLString({
		name: 'numFmt',
		attributes: attributes
	});
};

NumberFormats.prototype = _.merge({}, StylePart.prototype, {
	init: function init() {
		var _this = this;

		_.forEach(PREDEFINED, function (formatId, format) {
			_this.formatsByNames[format] = {
				formatId: formatId,
				format: format
			};
		});
	},

	canon: NumberFormats.canon,
	saveFormat: NumberFormats.saveFormat,
	merge: function merge(formatTo, formatFrom) {
		return formatFrom || formatTo;
	}
});

module.exports = NumberFormats;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./predeinedFormats":21,"./stylePart":23}],21:[function(require,module,exports){
'use strict';

var PREDEFINED = {
	date: 14, //mm-dd-yy
	time: 21 //h:mm:ss
};

module.exports = PREDEFINED;

},{}],22:[function(require,module,exports){
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
},{"../XMLString":2}],23:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

function StylePart(styles, saveName, formatName) {
	this.styles = styles;
	this.saveName = saveName;
	this.formatName = formatName;
	this.lastName = 1;
	this.lastId = 0;
	this.saveEmpty = true;
	this.formats = [];
	this.formatsByData = Object.create(null);
	this.formatsByNames = Object.create(null);
	this.predefined = {};
}

StylePart.prototype = {
	init: function init() {},
	add: function add(format, name, flags) {
		if (name && this.formatsByNames[name]) {
			var _canonFormat = this.canon(format, flags);
			var _stringFormat = _.isObject(_canonFormat) ? JSON.stringify(_canonFormat) : _canonFormat;

			if (_stringFormat !== this.formatsByNames[name].stringFormat) {
				this.addNew(_canonFormat, _stringFormat, name);
			}
			return name;
		}

		//first argument is format name
		if (!name && _.isString(format)) {
			if (this.formatsByNames[format]) {
				return format;
			} else if (this.predefined[format]) {
				return this.add(this.predefined[format], format);
			}
		}

		var canonFormat = this.canon(format, flags);
		var stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;
		var styleFormat = this.formatsByData[stringFormat];

		if (!styleFormat) {
			styleFormat = this.addNew(canonFormat, stringFormat, name);
		} else if (name && !this.formatsByNames[name]) {
			styleFormat.name = name;
			this.formatsByNames[name] = styleFormat;
		}
		return styleFormat.name;
	},
	addNew: function addNew(canonFormat, stringFormat, name) {
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
	},
	canon: function canon(format) {
		return format;
	},
	get: function get(format) {
		if (_.isString(format)) {
			var styleFormat = this.formatsByNames[format];

			return styleFormat ? styleFormat.format : format;
		}
		return format;
	},
	getId: function getId(name) {
		var styleFormat = this.formatsByNames[name];

		return styleFormat ? styleFormat.formatId : this.getPredefinedId(name);
	},
	getPredefinedId: function getPredefinedId(name) {
		if (this.predefined[name]) {
			return this.getId(this.add(this.predefined[name], name));
		}
		return null;
	},
	save: function save() {
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
	},
	saveCollectionExt: function saveCollectionExt() {},
	saveFormat: function saveFormat() {
		return '';
	}
};

module.exports = StylePart;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],24:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var NumberFormats = require('./numberFormats');
var Fonts = require('./fonts');
var Fills = require('./fills');
var Borders = require('./borders');
var alignment = require('./alignment');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.differentialformats.aspx
function TableElements(styles) {
	StylePart.call(this, styles, 'dxfs', 'tableElement');
}

TableElements.prototype = _.merge({}, StylePart.prototype, {
	canon: function canon(format) {
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
	},
	saveFormat: function saveFormat(format) {
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
});

module.exports = TableElements;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./alignment":14,"./borders":15,"./fills":17,"./fonts":18,"./numberFormats":20,"./stylePart":23}],25:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestylevalues.aspx
var ELEMENTS = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn', 'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe', 'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell'];
var SIZED_ELEMENTS = ['firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyles.aspx
function Tables(styles) {
	StylePart.call(this, styles, 'tableStyles', 'table');

	this.saveEmpty = false;
}

Tables.prototype = _.merge({}, StylePart.prototype, {
	canon: function canon(format) {
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
	},
	saveCollectionExt: function saveCollectionExt(attributes) {
		if (this.styles.defaultTableStyle) {
			attributes.push(['defaultTableStyle', this.styles.defaultTableStyle]);
		}
	},

	//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyleelement.aspx
	saveFormat: function saveFormat(format, styleFormat) {
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
});

module.exports = Tables;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./stylePart":23}],26:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;

var _require = require('../util'),
    canonColor = _require.canonColor;

var toXMLString = require('../XMLString');

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
},{"../XMLString":2,"../util":29}],27:[function(require,module,exports){
'use strict';

var Table = require('./table');

function createTable(outerWorksheet, common, config) {
	var table = new Table(config, common);

	var outerTable = {
		end: function end() {
			return outerWorksheet;
		},
		setReferenceRange: function setReferenceRange(beginCell, endCell) {
			table.setReferenceRange(beginCell, endCell);
			return this;
		},
		addTotalRow: function addTotalRow(totalRow) {
			table.addTotalRow(totalRow);
			return this;
		},
		setTheme: function setTheme(theme) {
			table.setTheme(theme);
			return this;
		},

		/**
   * Expects an object with the following properties:
   * caseSensitive (boolean)
   * dataRange
   * columnSort (assumes true)
   * sortDirection
   * sortRange (defaults to dataRange)
   */
		setSortState: function setSortState(state) {
			table.setSortState(state);
		}
	};
	return {
		outerTable: outerTable,
		table: table
	};
}

module.exports = createTable;

},{"./table":28}],28:[function(require,module,exports){
(function (global){
'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

function Table(config, common) {
	this.common = common;

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

var SUB_TOTAL_FUNCTIONS = ['average', 'countNums', 'count', 'max', 'min', 'stdDev', 'sum', 'var'];
var SUB_TOTAL_NUMS = [101, 102, 103, 104, 105, 107, 109, 110];

Table.prototype = {
	setReferenceRange: function setReferenceRange(beginCell, endCell) {
		this.beginCell = util.canonCell(beginCell);
		this.endCell = util.canonCell(endCell);
	},
	addTotalRow: function addTotalRow(totalRow) {
		this.totalRow = totalRow;
		this.totalsRowCount = 1;
	},
	setTheme: function setTheme(theme) {
		this.themeStyle = theme;
	},
	setSortState: function setSortState(state) {
		this.sortState = state;
	},
	prepare: function prepare(worksheetData) {
		if (this.totalRow) {
			var tableName = this.name;
			var beginCell = util.letterToPosition(this.beginCell);
			var endCell = util.letterToPosition(this.endCell);
			var firstRow = beginCell.y - 1;
			var firstColumn = beginCell.x - 1;
			var lastRow = endCell.y - 1;
			var headerRow = worksheetData[firstRow] || [];
			var totalRow = worksheetData[lastRow + 1];

			if (!totalRow) {
				totalRow = [];
				worksheetData[lastRow + 1] = totalRow;
			}

			_.forEach(this.totalRow, function (cell, i) {
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
		}
	},

	//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.table.aspx
	save: function save() {
		var attributes = [['id', this.tableId], ['name', this.name], ['displayName', this.displayName]];
		var children = [];
		var end = util.letterToPosition(this.endCell);
		var ref = this.beginCell + ':' + util.positionToLetter(end.x, end.y + this.totalsRowCount);

		attributes.push(['ref', ref]);
		attributes.push(['totalsRowCount', this.totalsRowCount]);
		attributes.push(['headerRowCount', this.headerRowCount]);

		children.push(this.saveAutoFilter());
		children.push(this.saveTableColumns());
		children.push(this.saveTableStyleInfo());

		return toXMLString({
			name: 'table',
			ns: 'spreadsheetml',
			attributes: attributes,
			children: children
		});
	},
	saveAutoFilter: function saveAutoFilter() {
		return toXMLString({
			name: 'autoFilter',
			attributes: ['ref', this.beginCell + ':' + this.endCell]
		});
	},
	saveTableColumns: function saveTableColumns() {
		var attributes = [['count', this.totalRow.length]];
		var children = _.map(this.totalRow, function (cell, index) {
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
	},
	saveTableStyleInfo: function saveTableStyleInfo() {
		var attributes = [['name', this.themeStyle]];
		var format = this.common.styles.tables.get(this.themeStyle);
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
};

module.exports = Table;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":29}],29:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;

var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var letterRefs = {};

function positionToLetter(x, y) {
	var result = letterRefs[x];

	if (!result) {
		var string = '';
		var num = x;
		var index = void 0;

		do {
			index = (num - 1) % 26;
			string = alphabet[index] + string;
			num = (num - (index + 1)) / 26;
		} while (num > 0);

		letterRefs[x] = string;
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

function pixelsToEMUs(pixels) {
	return Math.round(pixels * 914400 / 96);
}

function canonCell(cell) {
	if (_.isObject(cell)) {
		return positionToLetter(cell.c || 1, cell.r || 1);
	}
	return cell;
}

function canonColor(color) {
	return color[0] === '#' ? 'FF' + color.substr(1) : color;
}

module.exports = {
	positionToLetter: positionToLetter,
	letterToPosition: letterToPosition,
	pixelsToEMUs: pixelsToEMUs,
	canonCell: canonCell,
	canonColor: canonColor,

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
},{}],30:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var JSZip = typeof window !== "undefined" ? window['JSZip'] : typeof global !== "undefined" ? global['JSZip'] : null;
var util = require('./util');
var Common = require('./common');
var Relations = require('./relations');
var createWorksheet = require('./worksheet');
var toXMLString = require('./XMLString');

var maxWorksheetNameLength = 31;

// inner workbook
function Workbook(outerWorkbook) {
	this.outerWorkbook = outerWorkbook;
	this.common = new Common();
	this.styles = this.common.styles;
	this.images = this.common.images;

	this.relations = new Relations(this.common);
	this.relations.add(this.styles, 'stylesheet');
}

Workbook.prototype = {
	addWorksheet: function addWorksheet(config) {
		config = _.defaults(config, {
			name: this.common.getNewWorksheetDefaultName()
		});

		// Microsoft Excel (2007, 2013) do not allow worksheet names longer than 31 characters
		// if the worksheet name is longer, Excel displays an 'Excel found unreadable content...' popup when opening the file
		if (config.name.length > maxWorksheetNameLength) {
			throw 'Microsoft Excel requires work sheet names to be less than ' + (maxWorksheetNameLength + 1) + ' characters long, work sheet name "' + config.name + '" is ' + config.name.length + ' characters long';
		}

		var _createWorksheet = createWorksheet(this.outerWorkbook, this.common, config),
		    outerWorksheet = _createWorksheet.outerWorksheet,
		    worksheet = _createWorksheet.worksheet;

		this.common.addWorksheet(worksheet);
		this.relations.add(worksheet, 'worksheet');

		return outerWorksheet;
	},


	/**
  * Turns a workbook into a downloadable file.
  * options - options to modify how the zip is created. See http://stuk.github.io/jszip/#doc_generate_options
  */
	save: function save(options) {
		var zip = new JSZip();

		this.generateFiles(zip);
		return zip.generateAsync(_.defaults(options, {
			compression: 'DEFLATE',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			type: 'base64'
		}));
	},
	saveAsNodeStream: function saveAsNodeStream(options) {
		var zip = new JSZip();

		this.generateFiles(zip);
		return zip.generateNodeStream(_.defaults(options, {
			compression: 'DEFLATE'
		}));
	},
	generateFiles: function generateFiles(zip) {
		this.prepareWorksheets();

		this.saveWorksheets(zip);
		this.saveTables(zip);
		this.saveImages(zip);
		this.saveDrawings(zip);
		this.saveStyles(zip);
		this.saveStrings(zip);
		zip.file('[Content_Types].xml', this.createContentTypes());
		zip.file('_rels/.rels', this.createWorkbookRelationship());
		zip.file('xl/workbook.xml', this.saveWorkbook());
		zip.file('xl/_rels/workbook.xml.rels', this.relations.save());
	},
	saveWorkbook: function saveWorkbook() {
		return toXMLString({
			name: 'workbook',
			ns: 'spreadsheetml',
			attributes: [['xmlns:r', util.schemas.relationships]],
			children: [this.bookViewsXML(), this.sheetsXML(), this.definedNamesXML()]
		});
	},
	bookViewsXML: function bookViewsXML() {
		return toXMLString({
			name: 'bookViews',
			children: [toXMLString({
				name: 'workbookView',
				attributes: [['activeTab', this.common.getActiveWorksheetIndex()]]
			})]
		});
	},
	sheetsXML: function sheetsXML() {
		var _this = this;

		var children = this.common.worksheets.map(function (worksheet, index) {
			return toXMLString({
				name: 'sheet',
				attributes: [['name', worksheet.name], ['sheetId', index + 1], ['r:id', _this.relations.getId(worksheet)], ['state', worksheet.getState()]]
			});
		});

		return toXMLString({
			name: 'sheets',
			children: children
		});
	},
	definedNamesXML: function definedNamesXML() {
		var isPrintTitles = this.common.worksheets.some(function (worksheet) {
			return worksheet.isPrintTitle();
		});

		if (isPrintTitles) {
			var children = [];

			this.common.worksheets.forEach(function (worksheet, index) {
				var value = worksheet.savePrintTitle();

				if (value) {
					children.push(toXMLString({
						name: 'definedName',
						value: value,
						attributes: [['name', '_xlnm.Print_Titles'], ['localSheetId', index]]
					}));
				}
			});

			return toXMLString({
				name: 'definedNames',
				children: children
			});
		}
		return '';
	},
	prepareWorksheets: function prepareWorksheets() {
		this.common.worksheets.forEach(function (worksheet) {
			worksheet.prepare();
		});
	},
	saveWorksheets: function saveWorksheets(zip) {
		this.common.worksheets.forEach(function (worksheet) {
			zip.file(worksheet.path, worksheet.save());
			zip.file(worksheet.relationsPath, worksheet.relations.save());
		});
	},
	saveTables: function saveTables(zip) {
		this.common.tables.forEach(function (table) {
			zip.file(table.path, table.save());
		});
	},
	saveImages: function saveImages(zip) {
		_.forEach(this.images.getImages(), function (image) {
			zip.file(image.path, image.data, { base64: true, binary: true });
			image.data = null;
		});
		this.images.removeImages();
	},
	saveDrawings: function saveDrawings(zip) {
		this.common.drawings.forEach(function (drawing) {
			zip.file(drawing.path, drawing.save());
			zip.file(drawing.relationsPath, drawing.relations.save());
		});
	},
	saveStyles: function saveStyles(zip) {
		zip.file('xl/styles.xml', this.styles.save());
	},
	saveStrings: function saveStrings(zip) {
		if (this.common.strings.isStrings()) {
			this.relations.add(this.common.strings, 'sharedStrings');
			zip.file('xl/sharedStrings.xml', this.common.strings.save());
		}
	},
	createContentTypes: function createContentTypes() {
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
		if (this.common.strings.isStrings()) {
			children.push(toXMLString({
				name: 'Override',
				attributes: [['PartName', '/xl/sharedStrings.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml']]
			}));
		}
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/styles.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml']]
		}));

		this.common.worksheets.forEach(function (worksheet, index) {
			children.push(toXMLString({
				name: 'Override',
				attributes: [['PartName', '/xl/worksheets/sheet' + (index + 1) + '.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']]
			}));
		});
		this.common.tables.forEach(function (table, index) {
			children.push(toXMLString({
				name: 'Override',
				attributes: [['PartName', '/xl/tables/table' + (index + 1) + '.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']]
			}));
		});
		_.forEach(this.common.images.getExtensions(), function (contentType, extension) {
			children.push(toXMLString({
				name: 'Default',
				attributes: [['Extension', extension], ['ContentType', contentType]]
			}));
		});
		this.common.drawings.forEach(function (drawing, index) {
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
	},
	createWorkbookRelationship: function createWorkbookRelationship() {
		return toXMLString({
			name: 'Relationships',
			ns: 'relationshipPackage',
			children: [toXMLString({
				name: 'Relationship',
				attributes: [['Id', 'rId1'], ['Type', util.schemas.officeDocument], ['Target', 'xl/workbook.xml']]
			})]
		});
	}
};

module.exports = Workbook;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./common":4,"./relations":13,"./util":29,"./worksheet":33}],31:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var Drawings = require('../drawings');
var toXMLString = require('../XMLString');

function WorksheetDrawings(common, relations) {
	this.common = common;
	this.relations = relations;
	this.drawings = null;
}

WorksheetDrawings.prototype = {
	setImage: function setImage(image, config) {
		this.setDrawing(image, config, 'anchor');
	},
	setImageOneCell: function setImageOneCell(image, config) {
		this.setDrawing(image, config, 'oneCell');
	},
	setImageAbsolute: function setImageAbsolute(image, config) {
		this.setDrawing(image, config, 'absolute');
	},
	insert: function insert(colIndex, rowIndex, image) {
		if (image) {
			var cell = { c: colIndex + 1, r: rowIndex + 1 };

			if (typeof image === 'string' || image.data) {
				this.setDrawing(image, cell, 'anchor');
			} else {
				var config = image.config || {};
				config.cell = cell;

				this.setDrawing(image.image, config, 'anchor');
			}
		}
	},
	setDrawing: function setDrawing(image, config, anchorType) {
		if (!this.drawings) {
			this.drawings = new Drawings(this.common);

			this.common.addDrawings(this.drawings);
			this.relations.add(this.drawings, 'drawingRelationship');
		}

		var name = _.isObject(image) ? this.common.images.add(image.data, image.type) : image;

		this.drawings.add(name, config, anchorType);
	},
	save: function save() {
		if (this.drawings) {
			return toXMLString({
				name: 'drawing',
				attributes: [['r:id', this.relations.getId(this.drawings)]]
			});
		}
		return '';
	}
};

module.exports = WorksheetDrawings;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../drawings":10}],32:[function(require,module,exports){
'use strict';

var util = require('../util');
var toXMLString = require('../XMLString');

function Hyperlinks(common, relations) {
	this.common = common;
	this.relations = relations;
	this.hyperlinks = [];
}

Hyperlinks.prototype = {
	add: function add(hyperlink) {
		hyperlink.objectId = this.common.uniqueId('hyperlink');
		this.relations.add({
			objectId: hyperlink.objectId,
			target: hyperlink.location,
			targetMode: hyperlink.targetMode || 'External'
		}, 'hyperlink');
		this.hyperlinks.push(hyperlink);
	},
	insert: function insert(colIndex, rowIndex, hyperlink) {
		if (hyperlink) {
			var cell = { c: colIndex + 1, r: rowIndex + 1 };

			if (typeof hyperlink === 'string') {
				this.add({
					cell: cell,
					location: hyperlink
				});
			} else {
				this.add({
					cell: cell,
					location: hyperlink.location,
					targetMode: hyperlink.targetMode,
					tooltip: hyperlink.tooltip
				});
			}
		}
	},
	save: function save() {
		var _this = this;

		if (this.hyperlinks.length) {
			var children = this.hyperlinks.map(function (hyperlink) {
				var attributes = [['ref', util.canonCell(hyperlink.cell)], ['r:id', _this.relations.getId(hyperlink)]];

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
};

module.exports = Hyperlinks;

},{"../XMLString":2,"../util":29}],33:[function(require,module,exports){
'use strict';

var Worksheet = require('./worksheet');

function createWorksheet(outerWorkbook, common, config) {
	var outerWorksheet = {
		end: function end() {
			return outerWorkbook;
		},
		setActive: function setActive() {
			worksheet.common.setActiveWorksheet(worksheet);
			return this;
		},
		setVisible: function setVisible() {
			this.setState('visible');
			return this;
		},
		setHidden: function setHidden() {
			this.setState('hidden');
			return this;
		},

		/**
   * //http://www.datypic.com/sc/ooxml/t-ssml_ST_SheetState.html
   * @param state - visible | hidden | veryHidden
   */
		setState: function setState(state) {
			worksheet.setState(state);
			return this;
		},
		getState: function getState() {
			return worksheet.getState();
		},
		setRows: function setRows(startRow, rows) {
			worksheet.setRows(startRow, rows);
			return this;
		},
		setRow: function setRow(rowIndex, row) {
			worksheet.setRow(rowIndex, row);
			return this;
		},
		setColumns: function setColumns(startColumn, columns) {
			worksheet.setColumns(startColumn, columns);
			return this;
		},
		setColumn: function setColumn(columnIndex, column) {
			worksheet.setColumn(columnIndex, column);
			return this;
		},
		setData: function setData(offset, data) {
			worksheet.setData(offset, data);
			return this;
		},
		setAttribute: function setAttribute(name, value) {
			worksheet.sheetView.setAttribute(name, value);
			return this;
		},
		freeze: function freeze(col, row, cell, activePane) {
			worksheet.sheetView.freeze(col, row, cell, activePane);
			return this;
		},
		split: function split(x, y, cell, activePane) {
			worksheet.sheetView.split(x, y, cell, activePane);
			return this;
		},
		setHyperlink: function setHyperlink(hyperlink) {
			worksheet.hyperlinks.add(hyperlink);
			return this;
		},
		mergeCells: function mergeCells(cell1, cell2) {
			worksheet.mergedCells.mergeCells(cell1, cell2);
			return this;
		},
		setImage: function setImage(image, config) {
			worksheet.drawings.setImage(image, config);
			return this;
		},
		setImageOneCell: function setImageOneCell(image, config) {
			worksheet.drawings.setImageOneCell(image, config);
			return this;
		},
		setImageAbsolute: function setImageAbsolute(image, config) {
			worksheet.drawings.setImageAbsolute(image, config);
			return this;
		},
		addTable: function addTable(config) {
			return worksheet.tables.add(config);
		},
		setHeader: function setHeader(headers) {
			worksheet.setHeader(headers);
			return this;
		},
		setFooter: function setFooter(footers) {
			worksheet.setFooter(footers);
			return this;
		},
		setPageMargin: function setPageMargin(margin) {
			worksheet.setPageMargin(margin);
			return this;
		},
		setPageOrientation: function setPageOrientation(orientation) {
			worksheet.setPageOrientation(orientation);
			return this;
		},
		setPrintTitleTop: function setPrintTitleTop(params) {
			worksheet.setPrintTitleTop(params);
			return this;
		},
		setPrintTitleLeft: function setPrintTitleLeft(params) {
			worksheet.setPrintTitleLeft(params);
			return this;
		},
		addSparklineType: function addSparklineType(params, name) {
			return worksheet.sparklines.addType(params, name);
		},
		setDefaultSparklineType: function setDefaultSparklineType(params) {
			worksheet.sparklines.setDefaultType(params);
			return this;
		},
		addSparkline: function addSparkline(params) {
			worksheet.sparklines.add(params);
			return this;
		}
	};
	var worksheet = new Worksheet(outerWorksheet, common, config);

	return {
		outerWorksheet: outerWorksheet,
		worksheet: worksheet
	};
}

module.exports = createWorksheet;

},{"./worksheet":41}],34:[function(require,module,exports){
(function (global){
'use strict';

function _toConsumableArray(arr) { if (Array.isArray(arr)) { for (var i = 0, arr2 = Array(arr.length); i < arr.length; i++) { arr2[i] = arr[i]; } return arr2; } else { return Array.from(arr); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

function MergedCells(worksheet) {
	this.worksheet = worksheet;
	this.mergedCells = [];
}

MergedCells.prototype = {
	mergeCells: function mergeCells(cell1, cell2) {
		this.mergedCells.push([cell1, cell2]);
	},
	insert: function insert(dataRow, colIndex, rowIndex, colSpan, rowSpan, style) {
		if (colSpan) {
			dataRow = [].concat(_toConsumableArray(dataRow.slice(0, colIndex + 1)), _toConsumableArray(_.times(colSpan, function () {
				return { style: style, type: 'empty' };
			})), _toConsumableArray(dataRow.slice(colIndex + 1)));
		}
		if (rowSpan) {
			var data = this.worksheet.data;

			_.forEach(_.range(rowIndex + 1, rowIndex + 1 + rowSpan), function (index) {
				var row = data[index] || [];

				if (_.isArray(row)) {
					row = {
						data: row
					};
					data[index] = row;
				}

				row.inserts = row.inserts || [];
				_.forEach(_.range(colIndex, colIndex + colSpan + 1), function (index) {
					row.inserts[index] = { style: style };
				});
			});
		}
		return dataRow;
	},
	save: function save() {
		if (this.mergedCells.length) {
			var children = this.mergedCells.map(function (mergeCell) {
				return toXMLString({
					name: 'mergeCell',
					attributes: [['ref', util.canonCell(mergeCell[0]) + ':' + util.canonCell(mergeCell[1])]]
				});
			});

			return toXMLString({
				name: 'mergeCells',
				attributes: [['count', this.mergedCells.length]],
				children: children
			});
		}
		return '';
	}
};

module.exports = MergedCells;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":29}],35:[function(require,module,exports){
(function (global){
'use strict';

var _typeof = typeof Symbol === "function" && typeof Symbol.iterator === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof Symbol === "function" && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; };

function _toConsumableArray(arr) { if (Array.isArray(arr)) { for (var i = 0, arr2 = Array(arr.length); i < arr.length; i++) { arr2[i] = arr[i]; } return arr2; } else { return Array.from(arr); } }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;

var methods = {
	prepare: function prepare() {
		this.tables.prepare();
		this.prepareColumns();
		this.prepareRows();
		this.prepareData();
	},
	prepareColumns: function prepareColumns() {
		var _this = this;

		this.preparedColumns = this.columns.map(function (column) {
			if (column) {
				var preparedColumn = _.clone(column);

				if (column.style) {
					var style = _this.styles.addFormat(column.style);
					preparedColumn.style = style;

					var columnStyle = _this.styles.addFillOutFormat(style);
					preparedColumn.selfStyleId = _this.styles.getId(columnStyle);
				}
				return preparedColumn;
			}
			return undefined;
		});
		this.columns = null;
	},
	prepareRows: function prepareRows() {
		var _this2 = this;

		this.preparedRows = this.rows.map(function (row) {
			if (row) {
				var preparedRow = _.clone(row);

				if (row.style) {
					preparedRow.style = _this2.styles.addFormat(row.style);
				}
				return preparedRow;
			}
			return undefined;
		});
		this.rows = null;
	},
	prepareData: function prepareData() {
		this.maxX = 0;
		this.preparedData = [];
		for (var rowIndex = 0; rowIndex < this.data.length; rowIndex++) {
			var preparedDataRow = this.prepareDataRow(rowIndex);

			this.preparedData.push(preparedDataRow);
			this.maxX = Math.max(this.maxX, preparedDataRow.length);
		}
		this.maxY = this.preparedData.length;
		this.data = null;
	},
	prepareDataRow: function prepareDataRow(rowIndex) {
		var preparedDataRow = [];
		var row = this.preparedRows[rowIndex];
		var dataRow = this.data[rowIndex];

		if (dataRow) {
			var rowStyle = null;
			var skipColumnsStyle = false;
			var inserts = [];

			if (!_.isArray(dataRow)) {
				row = this.mergeDataRowToRow(row, dataRow);
				if (dataRow.inserts) {
					inserts = dataRow.inserts;
					dataRow = _.clone(dataRow.data);
				} else {
					dataRow = dataRow.data;
				}
			}
			if (row) {
				rowStyle = row.style || null;
				skipColumnsStyle = row.skipColumnsStyle;
			}
			dataRow = this.splitDataRow(row, dataRow, rowIndex);

			for (var colIndex = 0, dataIndex = 0; dataIndex < dataRow.length || colIndex < inserts.length; colIndex++) {
				var column = this.preparedColumns[colIndex];
				var columnStyle = !skipColumnsStyle && column ? column.style : null;

				var value = void 0;
				if (inserts[colIndex]) {
					value = { style: inserts[colIndex].style, type: 'empty' };
				} else {
					value = dataRow[dataIndex];
					dataIndex++;
				}

				var _readCellValue = this.readCellValue(value),
				    cellValue = _readCellValue.cellValue,
				    cellType = _readCellValue.cellType,
				    cellStyle = _readCellValue.cellStyle,
				    isObject = _readCellValue.isObject;

				var cellType2 = this.defineCellType(cellType, cellValue, row, column);
				var cellStyle2 = this.styles.getId(this.styles.merge(this.dataTimeStyle(cellType2), columnStyle, rowStyle, cellStyle));

				if (isObject) {
					this.hyperlinks.insert(colIndex, rowIndex, value.hyperlink);
					this.drawings.insert(colIndex, rowIndex, value.image);
					dataRow = this.mergeCells(dataRow, colIndex, rowIndex, value);
				}

				preparedDataRow[colIndex] = this.getPreparedCell(cellStyle2, cellType2, cellValue);
			}
		}

		if (row) {
			this.setRowStyleId(row);
			this.preparedRows[rowIndex] = row;
		}

		return preparedDataRow;
	},
	mergeDataRowToRow: function mergeDataRowToRow() {
		var row = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};
		var dataRow = arguments[1];

		row.height = dataRow.height || row.height;
		row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;
		row.type = dataRow.type || row.type;
		row.style = dataRow.style ? this.styles.addFormat(dataRow.style) : row.style;
		row.skipColumnsStyle = dataRow.skipColumnsStyle || row.skipColumnsStyle;

		return row;
	},
	splitDataRow: function splitDataRow() {
		var row = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};

		var _this3 = this,
		    _data;

		var dataRow = arguments[1];
		var rowIndex = arguments[2];

		var count = this.calcDataRowNestingCount(dataRow);

		if (count === 0) {
			return dataRow;
		}

		var newDataRows = _.times(count, function () {
			var result = _.clone(row);
			result.data = [];
			return result;
		});
		_.forEach(dataRow, function (value) {
			var list = void 0;
			var style = null;

			if (_.isArray(value)) {
				list = value;
			} else if (_.isObject(value) && !_.isDate(value) && _.isArray(value.value)) {
				list = value.value;
				style = value.style;
			}

			if (list) {
				_(list).initial().forEach(function (value, index) {
					newDataRows[index].data.push({ value: value, style: style });
				});

				var lastValue = { value: _.last(list), style: style };
				var listLength = list.length;
				var _value = listLength < count ? _this3.addRowspan(lastValue, count - listLength + 1) : lastValue;
				newDataRows[list.length - 1].data.push(_value);
			} else {
				newDataRows[0].data.push(_this3.addRowspan(value, count));
			}
		});
		(_data = this.data).splice.apply(_data, [rowIndex, 1].concat(_toConsumableArray(newDataRows)));

		return newDataRows[0].data;
	},
	calcDataRowNestingCount: function calcDataRowNestingCount(dataRow) {
		var count = 0;
		_.forEach(dataRow, function (value) {
			if (_.isArray(value)) {
				count = Math.max(value.length, count);
			} else if (_.isObject(value) && !_.isDate(value) && _.isArray(value.value)) {
				count = Math.max(value.value.length, count);
			}
		});
		return count;
	},
	addRowspan: function addRowspan(value, rowspan) {
		if (rowspan > 1) {
			if (_.isObject(value) && !_.isDate(value)) {
				value = _.clone(value);
				value.rowspan = rowspan;
				value.style = this.styles.merge({ vertical: 'top' }, value.style);
				return value;
			}
			return {
				value: value,
				rowspan: rowspan,
				style: { vertical: 'top' }
			};
		}
		return value;
	},
	readCellValue: function readCellValue(value) {
		var cellValue = void 0;
		var cellType = null;
		var cellStyle = null;
		var isObject = false;

		if (_.isDate(value)) {
			cellValue = value;
			cellType = 'dateOrTime';
		} else if (value && (typeof value === 'undefined' ? 'undefined' : _typeof(value)) === 'object') {
			isObject = true;

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
			} else if (_.isDate(value.value)) {
				cellValue = value.value;
				cellType = 'date';
			} else {
				cellValue = value.value;
				cellType = value.type;
			}
		} else {
			cellValue = value;
		}

		return {
			cellValue: cellValue,
			cellType: cellType,
			cellStyle: cellStyle,
			isObject: isObject
		};
	},
	mergeCells: function mergeCells(dataRow, colIndex, rowIndex, value) {
		if (value.colspan || value.rowspan) {
			var colSpan = (value.colspan || 1) - 1;
			var rowSpan = (value.rowspan || 1) - 1;

			if (colSpan || rowSpan) {
				this.mergedCells.mergeCells({ c: colIndex + 1, r: rowIndex + 1 }, { c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan });
				return this.mergedCells.insert(dataRow, colIndex, rowIndex, colSpan, rowSpan, value.style);
			}
		}
		return dataRow;
	},
	defineCellType: function defineCellType(cellType, cellValue, row, column) {
		if (cellType) {
			if (cellType === 'dateOrTime') {
				var type = row && row.type || column && column.type;

				return type === 'date' || type === 'time' ? type : 'date';
			}
			return cellType;
		} else if (row && row.type) {
			return row.type;
		} else if (column && column.type) {
			return column.type;
		} else if (typeof cellValue === 'number') {
			return 'number';
		} else if (typeof cellValue === 'string') {
			return 'string';
		}
	},
	dataTimeStyle: function dataTimeStyle(cellType) {
		return cellType === 'date' || cellType === 'time' ? cellType : null;
	},
	getPreparedCell: function getPreparedCell(styleId, cellType, cellValue) {
		var result = {
			styleId: styleId,
			value: null,
			formula: null,
			isString: false
		};

		if (cellType === 'string') {
			result.value = this.common.strings.add(cellValue);
			result.isString = true;
		} else if (cellType === 'date' || cellType === 'time') {
			var dateValue = _.isDate(cellValue) ? cellValue.valueOf() : cellValue;
			var date = 25569.0 + (dateValue - this.timezoneOffset) / (60 * 60 * 24 * 1000);

			if (_.isFinite(date)) {
				result.value = date;
			} else {
				result.value = this.common.strings.add(String(cellValue));
				result.isString = true;
			}
		} else if (cellType === 'formula') {
			result.formula = _.escape(cellValue);
		} else {
			result.value = cellValue;
		}

		return result;
	},
	setRowStyleId: function setRowStyleId(row) {
		if (row.style) {
			var rowStyle = this.styles.addFillOutFormat(row.style);
			row.selfStyleId = this.styles.getId(rowStyle);
		}
	}
};

module.exports = { methods: methods };

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{}],36:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

var methods = {
	/**
  * Expects an array length of three.
  * @param {Array} headers [left, center, right]
  * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.evenheader.aspx
  */
	setHeader: function setHeader(headers) {
		if (!_.isArray(headers)) {
			throw 'Invalid argument type - setHeader expects an array of three instructions';
		}
		this.headers = headers;
	},

	/**
  * Expects an array length of three.
  * @param {Array} footers [left, center, right]
  */
	setFooter: function setFooter(footers) {
		if (!_.isArray(footers)) {
			throw 'Invalid argument type - setFooter expects an array of three instructions';
		}
		this.footers = footers;
	},

	/**
  * Set page details in inches.
  */
	setPageMargin: function setPageMargin(margin) {
		this.margin = _.defaults(margin, {
			left: 0.7,
			right: 0.7,
			top: 0.75,
			bottom: 0.75,
			header: 0.3,
			footer: 0.3
		});
	},

	/**
  * http://www.datypic.com/sc/ooxml/t-ssml_ST_Orientation.html
  *
  * Can be one of 'portrait' or 'landscape'.
  *
  * @param {String} orientation
  */
	setPageOrientation: function setPageOrientation(orientation) {
		if (orientation === 'portrait' || orientation === 'landscape') {
			this.orientation = orientation;
		}
	},

	/**
  * Set rows to repeat for print
  *
  * @param {int|[int, int]} params - number of rows to repeat from the top | [first, last] repeat rows
  */
	setPrintTitleTop: function setPrintTitleTop(params) {
		this.printTitles = this.printTitles || {};

		if (_.isObject(params)) {
			this.printTitles.topFrom = params[0];
			this.printTitles.topTo = params[1];
		} else {
			this.printTitles.topFrom = 0;
			this.printTitles.topTo = params - 1;
		}
	},

	/**
  * Set columns to repeat for print
  *
  * @param {int|[int, int]} params - number of columns to repeat from the left | [first, last] repeat columns
  */
	setPrintTitleLeft: function setPrintTitleLeft(params) {
		this.printTitles = this.printTitles || {};

		if (_.isObject(params)) {
			this.printTitles.leftFrom = params[0];
			this.printTitles.leftTo = params[1];
		} else {
			this.printTitles.leftFrom = 0;
			this.printTitles.leftTo = params - 1;
		}
	},
	isPrintTitle: function isPrintTitle() {
		return this.printTitles && (this.printTitles.topTo >= 0 || this.printTitles.leftTo >= 0);
	},
	savePrintTitle: function savePrintTitle() {
		if (this.isPrintTitle()) {
			var printTitles = this.printTitles;
			var value = '';

			if (printTitles.topTo >= 0) {
				value = this.sheetName() + '!$' + (printTitles.topFrom + 1) + ':$' + (printTitles.topTo + 1);

				if (printTitles.leftTo >= 0) {
					value += ',';
				}
			}
			if (printTitles.leftTo >= 0) {
				value += this.sheetName() + '!$' + util.positionToLetter(printTitles.leftFrom + 1) + ':$' + util.positionToLetter(printTitles.leftTo + 1);
			}
			return value;
		}
		return '';
	},
	savePrint: function savePrint() {
		return this.savePageMargins() + this.savePageSetup() + this.saveHeaderFooter();
	},
	savePageMargins: function savePageMargins() {
		var margin = this.margin;

		if (margin) {
			return toXMLString({
				name: 'pageMargins',
				attributes: [['top', margin.top], ['bottom', margin.bottom], ['left', margin.left], ['right', margin.right], ['header', margin.header], ['footer', margin.footer]]
			});
		}
		return '';
	},
	savePageSetup: function savePageSetup() {
		if (this.orientation) {
			return toXMLString({
				name: 'pageSetup',
				attributes: [['orientation', this.orientation]]
			});
		}
		return '';
	},
	saveHeaderFooter: function saveHeaderFooter() {
		if (this.headers.length || this.footers.length) {
			var children = [];

			if (this.headers.length) {
				children.push(this.saveHeader());
			}
			if (this.footers.length) {
				children.push(this.saveFooter());
			}

			return toXMLString({
				name: 'headerFooter',
				children: children
			});
		}
		return '';
	},
	saveHeader: function saveHeader() {
		return toXMLString({
			name: 'oddHeader',
			value: compilePageDetailPackage(this.headers)
		});
	},
	saveFooter: function saveFooter() {
		return toXMLString({
			name: 'oddFooter',
			value: compilePageDetailPackage(this.footers)
		});
	}
};

function compilePageDetailPackage(data) {
	return ['&L', compilePageDetailPiece(data[0] || ''), '&C', compilePageDetailPiece(data[1] || ''), '&R', compilePageDetailPiece(data[2] || '')].join('');
}

function compilePageDetailPiece(data) {
	if (_.isString(data)) {
		return '&"-,Regular"' + data;
	} else if (_.isObject(data) && !_.isArray(data)) {
		var string = '&"' + (data.font || '-') + ',';

		if (data.bold && data.italic) {
			string += 'Bold Italic';
		} else if (data.bold) {
			string += 'Bold';
		} else if (data.italic) {
			string += 'Italic';
		} else {
			string += 'Regular';
		}
		string += '"';
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
			return result + compilePageDetailPiece(value);
		}, '');
	}
}

module.exports = { methods: methods };

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":29}],37:[function(require,module,exports){
(function (global){
'use strict';

var _createClass = function () { function defineProperties(target, props) { for (var i = 0; i < props.length; i++) { var descriptor = props[i]; descriptor.enumerable = descriptor.enumerable || false; descriptor.configurable = true; if ("value" in descriptor) descriptor.writable = true; Object.defineProperty(target, descriptor.key, descriptor); } } return function (Constructor, protoProps, staticProps) { if (protoProps) defineProperties(Constructor.prototype, protoProps); if (staticProps) defineProperties(Constructor, staticProps); return Constructor; }; }();

function _classCallCheck(instance, Constructor) { if (!(instance instanceof Constructor)) { throw new TypeError("Cannot call a class as a function"); } }

function _possibleConstructorReturn(self, call) { if (!self) { throw new ReferenceError("this hasn't been initialised - super() hasn't been called"); } return call && (typeof call === "object" || typeof call === "function") ? call : self; }

function _inherits(subClass, superClass) { if (typeof superClass !== "function" && superClass !== null) { throw new TypeError("Super expression must either be null or a function, not " + typeof superClass); } subClass.prototype = Object.create(superClass && superClass.prototype, { constructor: { value: subClass, enumerable: false, writable: true, configurable: true } }); if (superClass) Object.setPrototypeOf ? Object.setPrototypeOf(subClass, superClass) : subClass.__proto__ = superClass; }

var Readable = require('stream').Readable;
var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

var methods = {
	save: function save() {
		var canStream = !!Readable;

		if (canStream) {
			return new WorksheetStream({
				worksheet: this
			});
		}
		return saveBeforeRows(this) + saveData(this) + saveAfterRows(this);
	}
};

var WorksheetStream = function (_ref) {
	_inherits(WorksheetStream, _ref);

	function WorksheetStream(options) {
		_classCallCheck(this, WorksheetStream);

		var _this = _possibleConstructorReturn(this, (WorksheetStream.__proto__ || Object.getPrototypeOf(WorksheetStream)).call(this, options));

		_this.worksheet = options.worksheet;
		_this.status = 0;
		_this.index = 0;
		_this.len = _this.worksheet.preparedData.length;
		return _this;
	}

	_createClass(WorksheetStream, [{
		key: '_read',
		value: function _read(size) {
			var stop = false;

			if (this.status === 0) {
				stop = !this.push(saveBeforeRows(this.worksheet));

				this.status = 1;
			}

			if (this.status === 1) {
				while (this.index < this.len && !stop) {
					stop = !this.push(this.packChunk(size));
				}

				if (this.index === this.len) {
					this.status = 2;
				}
			}

			if (this.status === 2) {
				this.push(saveAfterRows(this.worksheet));
				this.push(null);
			}
		}
	}, {
		key: 'packChunk',
		value: function packChunk(size) {
			var worksheet = this.worksheet;
			var data = worksheet.preparedData;
			var preparedRows = worksheet.preparedRows;
			var s = '';

			while (this.index < this.len && s.length < size) {
				s += saveRow(data[this.index], preparedRows[this.index], this.index);
				data[this.index] = null;
				this.index++;
			}
			return s;
		}
	}]);

	return WorksheetStream;
}(Readable || null);

function saveBeforeRows(worksheet) {
	return util.xmlPrefix + '<worksheet xmlns="' + util.schemas.spreadsheetml + '" xmlns:r="' + util.schemas.relationships + '" xmlns:mc="' + util.schemas.markupCompat + '">' + saveDimension(worksheet.maxX, worksheet.maxY) + worksheet.sheetView.save() + saveColumns(worksheet.preparedColumns) + '<sheetData>';
}

function saveAfterRows(worksheet) {
	return '</sheetData>' +
	// 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
	// with Microsoft Excel (2007, 2013)
	worksheet.mergedCells.save() + worksheet.hyperlinks.save() + worksheet.savePrint() + worksheet.tables.save() +
	// the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
	// to issue with Microsoft Excel (2007, 2013)
	worksheet.drawings.save() + saveExtList(worksheet.sparklines) + '</worksheet>';
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
	if (dataRow) {
		var rowLen = dataRow.length;
		var rowChildren = new Array(rowLen);

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
		return '<row' + getRowAttributes(row, rowIndex) + '>' + rowChildren.join('') + '</row>';
	}
	return '<row' + getRowAttributes(row, rowIndex) + '></row>';
}

function getRowAttributes(row, rowIndex) {
	var attributes = ' r="' + (rowIndex + 1) + '"';

	if (row) {
		if (row.height !== undefined) {
			attributes += ' customHeight="1" ht="' + row.height + '"';
		}
		if (row.selfStyleId) {
			attributes += ' customFormat="1" s="' + row.selfStyleId + '"';
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
		var children = columns.map(function (column, index) {
			column = column || {};

			var attributes = [['min', column.min || index + 1], ['max', column.max || index + 1]];

			if (column.hidden) {
				attributes.push(['hidden', 1]);
			}
			if (column.bestFit) {
				attributes.push(['bestFit', 1]);
			}
			if (column.width) {
				attributes.push(['customWidth', 1]);
				attributes.push(['width', column.width]);
			} else {
				attributes.push(['width', 9.140625]);
			}
			if (column.selfStyleId) {
				attributes.push(['style', column.selfStyleId]);
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

function saveExtList(sparklines) {
	var sparklinesStr = sparklines.save();

	if (sparklinesStr) {
		return toXMLString({
			name: 'extLst',
			children: [sparklinesStr]
		});
	}
	return '';
}

module.exports = { methods: methods };

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":29,"stream":1}],38:[function(require,module,exports){
(function (global){
'use strict';

/**
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sheetview.aspx
 */

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

function SheetView(config) {
	var _this = this;

	this.pane = null;
	this.attributes = {
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
		} else if (_this.attributes[name]) {
			_this.attributes[name].value = value;
		}
	});
}

SheetView.prototype = {
	setAttribute: function setAttribute(name, value) {
		if (this.attributes[name]) {
			this.attributes[name].value = value;
		}
	},

	/**
  * Add froze pane
  * @param col - column number: 0, 1, 2 ...
  * @param row - row number: 0, 1, 2 ...
  * @param cell? - 'A1' | {c: 1, r: 1}
  * @param activePane? - topLeft | topRight | bottomLeft | bottomRight
  */
	freeze: function freeze(col, row, cell, activePane) {
		this.pane = {
			state: 'frozen',
			xSplit: col,
			ySplit: row,
			topLeftCell: util.canonCell(cell) || util.positionToLetter(col + 1, row + 1),
			activePane: activePane || 'bottomRight'
		};
	},

	/**
  * Add split pane
  * @param x - Horizontal position of the split, in points; 0 (zero) if none
  * @param y - Vertical position of the split, in points; 0 (zero) if none
  * @param cell? - 'A1' | {c: 1, r: 1}
  * @param activePane? - topLeft | topRight | bottomLeft | bottomRight
  */
	split: function split(x, y, cell, activePane) {
		this.pane = {
			state: 'split',
			xSplit: x * 20,
			ySplit: y * 20,
			topLeftCell: util.canonCell(cell) || 'A1',
			activePane: activePane || 'bottomRight'
		};
	},
	save: function save() {
		var attributes = [['workbookViewId', 0]];

		_.forEach(this.attributes, function (attr, name) {
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
				children: [this.savePane()]
			})]
		});
	},
	savePane: function savePane() {
		var pane = this.pane;

		if (pane) {
			return toXMLString({
				name: 'pane',
				attributes: [['state', pane.state], ['xSplit', pane.xSplit], ['ySplit', pane.ySplit], ['topLeftCell', pane.topLeftCell], ['activePane', pane.activePane]]
			});
		}
		return '';
	}
};

module.exports = SheetView;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":29}],39:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

var defaultTypeAttrs = {
	displayEmptyCellsAs: 'gap'
};

var defaultTypeColors = {
	color: 'FF376092',
	colorNegative: 'FFD00000',
	colorAxis: 'FF000000',
	colorMarkers: 'FFD00000',
	colorFirst: 'FFD00000',
	colorLast: 'FFD00000',
	colorHigh: 'FFD00000',
	colorLow: 'FFD00000'
};

function Sparklines(worksheet) {
	this.worksheet = worksheet;
	this.lines = Object.create(null);
	this.defaultLines = {
		type: null,
		lines: []
	};
	this.typesByData = Object.create(null);
	this.lastId = 1;
}

Sparklines.prototype = {
	addType: function addType(params, name) {
		var type = this.canonType(params);
		var stringType = JSON.stringify(type);

		if (!name && this.typesByData[stringType]) {
			return this.typesByData[stringType];
		}
		return this.addNew(type, stringType, name);
	},
	setDefaultType: function setDefaultType(params) {
		this.defaultLines.type = this.canonType(params);
	},
	add: function add() {
		var _ref = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {},
		    cell = _ref.cell,
		    beginCell = _ref.beginCell,
		    endCell = _ref.endCell,
		    type = _ref.type;

		var line = {
			cell: util.canonCell(cell),
			beginCell: util.canonCell(beginCell),
			endCell: util.canonCell(endCell)
		};

		if (type) {
			this.lines(type).lines.push(line);
		} else {
			this.defaultLines.lines.push(line);
		}
	},
	insert: function insert() {},
	canonType: function canonType() {
		var params = arguments.length > 0 && arguments[0] !== undefined ? arguments[0] : {};

		var type = {};

		_.keys(defaultTypeAttrs).forEach(function (property) {
			type[property] = params[property] || defaultTypeAttrs[property];
		});
		_.keys(defaultTypeColors).forEach(function (property) {
			type[property] = util.canonColor(params[property] || defaultTypeColors[property]);
		});
		return type;
	},
	addNew: function addNew(type, stringType, name) {
		name = name || 'Sparklines' + this.lastId++;

		this.typesByData[stringType] = name;
		this.lines(type).type = type;

		return name;
	},
	lines: function lines(type) {
		this.lines[type] = this.lines[type] || {
			type: null,
			lines: []
		};
		return this.lines[type];
	},
	save: function save() {
		var _this = this;

		var children = [];

		if (this.defaultLines.lines.length) {
			var type = this.defaultLines.type || this.canonType();

			children.push(this.saveType(type, this.defaultLines.lines));
		}

		_.forEach(this.lines, function (_ref2) {
			var type = _ref2.type,
			    lines = _ref2.lines;

			if (type && lines.length) {
				children.push(_this.saveType(type, lines));
			}
		});

		if (children.length) {
			return toXMLString({
				name: 'ext',
				attributes: [['uri', '{05C60535-1F16-4fd2-B633-F4F36F0B64E0}'], ['xmlns:x14', 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/main']],
				children: [toXMLString({
					name: 'x14:sparklineGroups',
					attributes: [['xmlns:xm', 'http://schemas.microsoft.com/office/excel/2006/main']],
					children: children
				})]
			});
		}
		return '';
	},
	saveType: function saveType(type, lines) {
		var attributes = [];
		var children = [];

		_.keys(defaultTypeAttrs).forEach(function (property) {
			attributes.push([property, type[property]]);
		});
		_.keys(defaultTypeColors).forEach(function (property) {
			var name = property === 'color' ? 'colorSeries' : property;
			var value = type[property];

			children.push('<x14:' + name + ' rgb="' + value + '"/>');
		});
		children.push(toXMLString({
			name: 'x14:sparklines',
			children: this.saveLines(lines)
		}));

		return toXMLString({
			name: 'x14:sparklineGroup',
			attributes: attributes,
			children: children
		});
	},
	saveLines: function saveLines(lines) {
		var _this2 = this;

		return lines.map(function (line) {
			return toXMLString({
				name: 'x14:sparkline',
				children: ['<xm:f>' + _this2.worksheet.sheetName() + '!' + line.beginCell + ':' + line.endCell + '</xm:f>', '<xm:sqref>' + line.cell + '</xm:sqref>']
			});
		});
	}
};

module.exports = Sparklines;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":29}],40:[function(require,module,exports){
'use strict';

var createTable = require('../table');
var toXMLString = require('../XMLString');

function Tables(worksheet, common, relations) {
	this.worksheet = worksheet;
	this.common = common;
	this.relations = relations;
	this.tables = [];
}

Tables.prototype = {
	add: function add(config) {
		var _createTable = createTable(this.worksheet.outerWorksheet, this.common, config),
		    outerTable = _createTable.outerTable,
		    table = _createTable.table;

		this.common.addTable(table);
		this.relations.add(table, 'table');
		this.tables.push(table);

		return outerTable;
	},
	prepare: function prepare() {
		var _this = this;

		this.tables.forEach(function (table) {
			table.prepare(_this.worksheet.data);
		});
	},
	save: function save() {
		var _this2 = this;

		if (this.tables.length) {
			var children = this.tables.map(function (table) {
				return toXMLString({
					name: 'tablePart',
					attributes: [['r:id', _this2.relations.getId(table)]]
				});
			});

			return toXMLString({
				name: 'tableParts',
				attributes: [['count', this.tables.length]],
				children: children
			});
		}
		return '';
	}
};

module.exports = Tables;

},{"../XMLString":2,"../table":27}],41:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var Tables = require('./tables');
var WorksheetDrawings = require('./drawing');
var Hyperlinks = require('./hyperlinks');
var MergedCells = require('./mergedCells');
var Sparklines = require('./sparklines');
var SheetView = require('./sheetView');
var print = require('./print');
var prepareSave = require('./prepareSave');
var save = require('./save');
var Relations = require('../relations');

function Worksheet(outerWorksheet, common) {
	var config = arguments.length > 2 && arguments[2] !== undefined ? arguments[2] : {};

	this.outerWorksheet = outerWorksheet;
	this.common = common;
	this.styles = this.common.styles;

	this.objectId = this.common.uniqueId('Worksheet');
	this.data = [];
	this.columns = [];
	this.rows = [];

	this.headers = [];
	this.footers = [];

	this.name = config.name;
	this.state = config.state || 'visible';
	this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
	this.relations = new Relations(this.common);

	this.tables = new Tables(this, this.common, this.relations);
	this.drawings = new WorksheetDrawings(this.common, this.relations);
	this.hyperlinks = new Hyperlinks(this.common, this.relations);
	this.mergedCells = new MergedCells(this);
	this.sparklines = new Sparklines(this);
	this.sheetView = new SheetView(config);
}

Worksheet.prototype = _.assign({
	setRows: function setRows(startRow, rows) {
		var _this = this;

		if (!rows) {
			rows = startRow;
			startRow = 0;
		} else {
			--startRow;
		}
		rows.forEach(function (row, i) {
			_this.rows[startRow + i] = row;
		});
	},
	setRow: function setRow(rowIndex, row) {
		this.rows[--rowIndex] = row;
	},

	/**
  * http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.aspx
  */
	setColumns: function setColumns(startColumn, columns) {
		var _this2 = this;

		if (!columns) {
			columns = startColumn;
			startColumn = 0;
		} else {
			--startColumn;
		}
		columns.forEach(function (column, i) {
			_this2.columns[startColumn + i] = column;
		});
	},
	setColumn: function setColumn(columnIndex, column) {
		this.columns[--columnIndex] = column;
	},
	setData: function setData(offset, data) {
		var _this3 = this;

		var startRow = this.data.length;

		if (!data) {
			data = offset;
		} else {
			startRow += offset;
		}
		data.forEach(function (row, i) {
			_this3.data[startRow + i] = row;
		});
	},
	setState: function setState(state) {
		this.state = state;
	},
	getState: function getState() {
		return this.state;
	},
	sheetName: function sheetName() {
		return '\'' + this.name + '\'';
	}
}, print.methods, prepareSave.methods, save.methods);

module.exports = Worksheet;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../relations":13,"./drawing":31,"./hyperlinks":32,"./mergedCells":34,"./prepareSave":35,"./print":36,"./save":37,"./sheetView":38,"./sparklines":39,"./tables":40}]},{},[12])(12)
});