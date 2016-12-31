/*
excelWriter - Javascript module for generating Excel files
<https://github.com/anfilat/excelWriter.js>
(c) 2016 Andrey Filatkin
Dual licenced under the MIT license or GPLv3
*/
(function(f){if(typeof exports==="object"&&typeof module!=="undefined"){module.exports=f()}else if(typeof define==="function"&&define.amd){define([],f)}else{var g;if(typeof window!=="undefined"){g=window}else if(typeof global!=="undefined"){g=global}else if(typeof self!=="undefined"){g=self}else{g=this}g.excelWriter = f()}})(function(){var define,module,exports;return (function e(t,n,r){function s(o,u){if(!n[o]){if(!t[o]){var a=typeof require=="function"&&require;if(!u&&a)return a(o,!0);if(i)return i(o,!0);var f=new Error("Cannot find module '"+o+"'");throw f.code="MODULE_NOT_FOUND",f}var l=n[o]={exports:{}};t[o][0].call(l.exports,function(e){var n=t[o][1][e];return s(n?n:e)},l,l.exports,e,t,n,r)}return n[o].exports}var i=typeof require=="function"&&require;for(var o=0;o<r.length;o++)s(r[o]);return s})({1:[function(require,module,exports){

},{}],2:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
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
	var name = config.name;
	var string = '<' + name;
	var content = '';
	var attr;
	var i, l;

	if (config.ns) {
		string = util.xmlPrefix + string +	' xmlns="' + util.schemas[config.ns] + '"';
	}
	if (config.attributes) {
		for (i = 0, l = config.attributes.length; i < l; i++) {
			attr = config.attributes[i];

			string += ' ' + attr[0] + '="' + _.escape(attr[1]) + '"';
		}
	}
	if (!_.isUndefined(config.value)) {
		content += _.escape(config.value);
	}
	if (config.children) {
		content += config.children.join('');
	}

	if (content) {
		string += '>' + content + '</' + name + '>';
	} else {
		string += '/>';
	}

	return string;
}

module.exports = toXMLString;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./util":27}],3:[function(require,module,exports){
'use strict';

module.exports = {
	init: function () {
		this._images = Object.create(null);
		this._imageByNames = Object.create(null);
		this._extensions = Object.create(null);
	},
	methods: {
		addImage: function (data, type, name) {
			var image = this._images[data];

			if (!image) {
				type = type || '';

				var contentType;
				var id = this.uniqueIdSeparated('image').id;
				var path = 'xl/media/image' + id + '.' + type;

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
				this.addPath(image, '/' + path);
				this._images[data] = image;
				this._imageByNames[name] = image;
				this._extensions[type] = contentType;
			} else if (name && !this._imageByNames[name]) {
				image.name = name;
				this._imageByNames[name] = image;
			}
			return image.name;
		},
		getImage: function (name) {
			return this._imageByNames[name];
		},
		getImages: function () {
			return this._images;
		},
		removeImages: function () {
			this._images = null;
			this._imageByNames = null;
		},
		getExtensions: function () {
			return this._extensions;
		}
	}
};

},{}],4:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var paths = require('./paths');
var images = require('./images');
var SharedStrings = require('./sharedStrings');
var Styles = require('../styles');

function Common() {
	paths.init.call(this);
	images.init.call(this);

	this.idSpaces = Object.create(null);

	this.sharedStrings = new SharedStrings(this);
	this.addPath(this.sharedStrings, 'sharedStrings.xml');

	this.styles = new Styles(this);
	this.addPath(this.styles, 'styles.xml');

	this.worksheets = [];
	this.tables = [];
	this.drawings = [];
}

_.assign(Common.prototype, paths.methods);
_.assign(Common.prototype, images.methods);

Common.prototype.uniqueId = function (space) {
	if (!this.idSpaces[space]) {
		this.idSpaces[space] = 1;
	}
	return space + this.idSpaces[space]++;
};

Common.prototype.uniqueIdSeparated = function (space) {
	if (!this.idSpaces[space]) {
		this.idSpaces[space] = 1;
	}
	return {
		space: space,
		id: this.idSpaces[space]++
	};
};

Common.prototype.addString = function (string) {
	return this.sharedStrings.add(string);
};

Common.prototype.addWorksheet = function (worksheet) {
	var index = this.worksheets.length + 1;
	var path = 'worksheets/sheet' + index + '.xml';
	var relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

	worksheet.path = 'xl/' + path;
	worksheet.relationsPath = relationsPath;
	this.worksheets.push(worksheet);
	this.addPath(worksheet, path);
};

Common.prototype.getNewWorksheetDefaultName = function () {
	return 'Sheet ' + (this.worksheets.length + 1);
};

Common.prototype.setActiveWorksheet = function (worksheet) {
	this.activeWorksheet = worksheet;
};

Common.prototype.addTable = function (table) {
	var index = this.tables.length + 1;
	var path = 'xl/tables/table' + index + '.xml';

	table.path = path;
	this.tables.push(table);
	this.addPath(table, '/' + path);
};

Common.prototype.addDrawings = function (drawings) {
	var index = this.drawings.length + 1;
	var path = 'xl/drawings/drawing' + index + '.xml';
	var relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

	drawings.path = path;
	drawings.relationsPath = relationsPath;
	this.drawings.push(drawings);
	this.addPath(drawings, '/' + path);
};

module.exports = Common;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../styles":19,"./images":3,"./paths":5,"./sharedStrings":6}],5:[function(require,module,exports){
'use strict';

module.exports = {
	init: function () {
		this._paths = Object.create(null);
	},
	methods: {
		addPath: function (object, path) {
			this._paths[object.objectId] = path;
		},
		getPath: function (object) {
			return this._paths[object.objectId];
		}
	}
};

},{}],6:[function(require,module,exports){
(function (global){
'use strict';

var Readable = require('stream').Readable;
var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('../util');

var spaceRE = /^\s|\s$/;

function SharedStrings(common) {
	this.objectId = common.uniqueId('SharedStrings');
	this._strings = Object.create(null);
	this._stringArray = [];
}

SharedStrings.prototype.add = function (string) {
	var stringId = this._strings[string];

	if (stringId === undefined) {
		stringId = this._stringArray.length;

		this._strings[string] = stringId;
		this._stringArray[stringId] = string;
	}

	return stringId;
};

SharedStrings.prototype.isEmpty = function () {
	return this._stringArray.length === 0;
};

SharedStrings.prototype.export = function (canStream) {
	this._strings = null;

	if (canStream) {
		return new SharedStringsStream({
			strings: this._stringArray
		});
	} else {
		var len = this._stringArray.length;
		var children = _.map(this._stringArray, function (string) {
			var str = _.escape(string);

			if (spaceRE.test(str)) {
				return '<si><t xml:space="preserve">' + str + '</t></si>';
			}
			return '<si><t>' + str + '</t></si>';
		});

		return getXMLBegin(len) + children.join('') + getXMLEnd();
	}
};

function SharedStringsStream(options) {
	Readable.call(this, options);

	this.strings = options.strings;
	this.status = 0;
}

util.inherits(SharedStringsStream, Readable || {});

SharedStringsStream.prototype._read = function (size) {
	var stop = false;

	if (this.status === 0) {
		stop = !this.push(getXMLBegin(this.strings.length));

		this.status = 1;
		this.index = 0;
		this.len = this.strings.length;
	}

	if (this.status === 1) {
		var s = '';
		var str;

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
};

function getXMLBegin(length) {
	return util.xmlPrefix + '<sst xmlns="' + util.schemas.spreadsheetml +
		'" count="' + length + '" uniqueCount="' + length + '">';
}

function getXMLEnd() {
	return '</sst>';
}

module.exports = SharedStrings;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../util":27,"stream":1}],7:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('../util');
var toXMLString = require('../XMLString');

function Anchor(config) {
	var coord;
	var x, y;

	if (_.isObject(config)) {
		if (_.has(config, 'cell')) {
			if (_.isObject(config.cell)) {
				coord = {x: config.cell.c || 1, y: config.cell.r || 1};
			} else {
				coord = util.letterToPosition(config.cell || '');
			}
		} else {
			coord = {x: config.c || 1, y: config.r || 1};
		}
	} else {
		coord = util.letterToPosition(config || '');
		config = {};
	}
	x = coord.x - 1;
	y = coord.y - 1;

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

Anchor.prototype.exportWithContent = function (content) {
	return toXMLString({
		name: 'xdr:twoCellAnchor',
		children: [
			toXMLString({
				name: 'xdr:from',
				children: [
					toXMLString({
						name: 'xdr:col',
						value: this.from.x
					}),
					toXMLString({
						name: 'xdr:colOff',
						value: this.from.xOff
					}),
					toXMLString({
						name: 'xdr:row',
						value: this.from.y
					}),
					toXMLString({
						name: 'xdr:rowOff',
						value: this.from.yOff
					})
				]
			}),
			toXMLString({
				name: 'xdr:to',
				children: [
					toXMLString({
						name: 'xdr:col',
						value: this.to.x
					}),
					toXMLString({
						name: 'xdr:colOff',
						value: this.to.xOff
					}),
					toXMLString({
						name: 'xdr:row',
						value: this.to.y
					}),
					toXMLString({
						name: 'xdr:rowOff',
						value: this.to.yOff
					})
				]
			}),
			content,
			toXMLString({
				name: 'xdr:clientData'
			})
		]
	});
};

module.exports = Anchor;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27}],8:[function(require,module,exports){
'use strict';

var util = require('../util');
var toXMLString = require('../XMLString');

function AnchorAbsolute(config) {
	config = config || {};

	this.x = util.pixelsToEMUs(config.left || 0);
	this.y = util.pixelsToEMUs(config.top || 0);
	this.width = util.pixelsToEMUs(config.width || 0);
	this.height = util.pixelsToEMUs(config.height || 0);
}

AnchorAbsolute.prototype.exportWithContent = function (content) {
	return toXMLString({
		name: 'xdr:absoluteAnchor',
		children: [
			toXMLString({
				name: 'xdr:pos',
				attributes: [
					['x', this.x],
					['y', this.y]
				]
			}),
			toXMLString({
				name: 'xdr:ext',
				attributes: [
					['cx', this.width],
					['cy', this.height]
				]
			}),
			content,
			toXMLString({
				name: 'xdr:clientData'
			})
		]
	});
};

module.exports = AnchorAbsolute;

},{"../XMLString":2,"../util":27}],9:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('../util');
var toXMLString = require('../XMLString');

function AnchorOneCell(config) {
	var coord;

	if (_.isObject(config)) {
		if (_.has(config, 'cell')) {
			if (_.isObject(config.cell)) {
				coord = {x: config.cell.c || 1, y: config.cell.r || 1};
			} else {
				coord = util.letterToPosition(config.cell || '');
			}
		} else {
			coord = {x: config.c || 1, y: config.r || 1};
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

AnchorOneCell.prototype.exportWithContent = function (content) {
	return toXMLString({
		name: 'xdr:oneCellAnchor',
		children: [
			toXMLString({
				name: 'xdr:from',
				children: [
					toXMLString({
						name: 'xdr:col',
						value: this.x
					}),
					toXMLString({
						name: 'xdr:colOff',
						value: this.xOff
					}),
					toXMLString({
						name: 'xdr:row',
						value: this.y
					}),
					toXMLString({
						name: 'xdr:rowOff',
						value: this.yOff
					})
				]
			}),
			toXMLString({
				name: 'xdr:ext',
				attributes: [
					['cx', this.width],
					['cy', this.height]
				]
			}),
			content,
			toXMLString({
				name: 'xdr:clientData'
			})
		]
	});
};

module.exports = AnchorOneCell;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27}],10:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
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

Drawings.prototype.addImage = function (name, config, anchorType) {
	var image = this.common.getImage(name);
	var imageRelationId = this.relations.addRelation(image, 'image');
	var picture = new Picture(this.common, {
		image: image,
		imageRelationId: imageRelationId,
		config: config,
		anchorType: anchorType
	});

	this.drawings.push(picture);
};

Drawings.prototype.export = function () {
	var attributes = [
		['xmlns:a', util.schemas.drawing],
		['xmlns:r', util.schemas.relationships],
		['xmlns:xdr', util.schemas.spreadsheetDrawing]
	];
	var children = _.map(this.drawings, function (picture) {
		return picture.export();
	});

	return toXMLString({
		name: 'xdr:wsDr',
		ns: 'spreadsheetDrawing',
		attributes: attributes,
		children: children
	});
};

module.exports = Drawings;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../relations":13,"../util":27,"./picture":11}],11:[function(require,module,exports){
'use strict';

var util = require('../util');
var toXMLString = require('../XMLString');
var Anchor = require('./anchor');
var AnchorOneCell = require('./anchorOneCell');
var AnchorAbsolute = require('./anchorAbsolute');

function Picture(common, config) {
	this.pictureId = common.uniqueIdSeparated('Picture').id;
	this.image = config.image;
	this.imageRelationId = config.imageRelationId;
	this.createAnchor(config.anchorType, config.config);
}

/**
 *
 * @param {String} type Can be 'anchor', 'oneCell' or 'absolute'.
 * @param {Object} config Shorthand - pass the created anchor coords that can normally be used to construct it.
 */
Picture.prototype.createAnchor = function (type, config) {
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
};

Picture.prototype.export = function () {
	var picture = toXMLString({
		name: 'xdr:pic',
		children: [
			toXMLString({
				name: 'xdr:nvPicPr',
				children: [
					toXMLString({
						name: 'xdr:cNvPr',
						attributes: [
							['id', this.pictureId],
							['name', this.image.name]
						]
					}),
					toXMLString({
						name: 'xdr:cNvPicPr',
						children: [
							toXMLString({
								name: 'a:picLocks',
								attributes: [
									['noChangeAspect', '1'],
									['noChangeArrowheads', '1']
								]
							})
						]
					})
				]
			}),
			toXMLString({
				name: 'xdr:blipFill',
				children: [
					toXMLString({
						name: 'a:blip',
						attributes: [
							['xmlns:r', util.schemas.relationships],
							['r:embed', this.imageRelationId]
						]
					}),
					toXMLString({
						name: 'a:srcRect'
					}),
					toXMLString({
						name: 'a:stretch',
						children: [
							toXMLString({
								name: 'a:fillRect'
							})
						]
					})
				]
			}),
			toXMLString({
				name: 'xdr:spPr',
				attributes: [
					['bwMode', 'auto']
				],
				children: [
					toXMLString({
						name: 'a:xfrm'
					}),
					toXMLString({
						name: 'a:prstGeom',
						attributes: [
							['prst', 'rect']
						]
					})
				]
			})
		]
	});

	return this.anchor.exportWithContent(picture);
};

module.exports = Picture;

},{"../XMLString":2,"../util":27,"./anchor":7,"./anchorAbsolute":8,"./anchorOneCell":9}],12:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var JSZip = (typeof window !== "undefined" ? window['JSZip'] : typeof global !== "undefined" ? global['JSZip'] : null);
var Workbook = require('./workbook');

var excelWriter = {
	createWorkbook: function () {
		return new Workbook();
	},

	/**
	 * Turns a workbook into a downloadable file.
	 * @param {Workbook} workbook - The workbook that is being converted
	 * @param {Object?} options - options to modify how the zip is created. See http://stuk.github.io/jszip/#doc_generate_options
	 */
	save: function (workbook, options) {
		var zip = new JSZip();
		var canStream = JSZip.support.nodestream;

		workbook._generateFiles(zip, canStream);
		return zip.generateAsync(_.defaults(options, {
			compression: 'DEFLATE',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			type: 'base64'
		}));
	},

	saveAsNodeStream: function (workbook, options) {
		var zip = new JSZip();
		var canStream = JSZip.support.nodestream;

		workbook._generateFiles(zip, canStream);
		return zip.generateNodeStream(_.defaults(options, {
			compression: 'DEFLATE'
		}));
	}
};

module.exports = excelWriter;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./workbook":28}],13:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('./util');
var toXMLString = require('./XMLString');

function RelationshipManager(common) {
	this.common = common;

	this.relations = Object.create(null);
	this.lastId = 1;
}

RelationshipManager.prototype.addRelation = function (object, type) {
	var relation = this.relations[object.objectId];
	var relationId;

	if (relation) {
		relationId = relation.relationId;
	} else {
		relationId = 'rId' + this.lastId++;

		this.relations[object.objectId] = {
			relationId: relationId,
			schema: util.schemas[type],
			object: object
		};
	}
	return relationId;
};

RelationshipManager.prototype.getRelationshipId = function (object) {
	var relation = this.relations[object.objectId];

	return relation ? relation.relationId : null;
};

RelationshipManager.prototype.export = function () {
	var common = this.common;
	var children = _.map(this.relations, function (relation) {
		var attributes = [
			['Id', relation.relationId],
			['Type', relation.schema],
			['Target', relation.object.target || common.getPath(relation.object)]
		];

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
};

module.exports = RelationshipManager;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./util":27}],14:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
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

function exportAlignment(format) {
	return toXMLString({
		name: 'alignment',
		attributes: _.toPairs(format)
	});
}

module.exports = {
	canon: canon,
	merge: merge,
	exportFormat: exportAlignment
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],15:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var StylePart = require('./stylePart');
var util = require('../util');
var formatUtils = require('./utils');
var toXMLString = require('../XMLString');

var BORDERS = ['left', 'right', 'top', 'bottom', 'diagonal'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borders.aspx
function Borders(styles) {
	StylePart.call(this, styles, 'borders', 'border');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Borders, StylePart);

Borders.canon = function (format) {
	var result = {};

	_.forEach(BORDERS, function (name) {
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
	return result;
};

Borders.exportFormat = function (format) {
	var children = _.map(BORDERS, function (name) {
		var border = format[name];
		var attributes;
		var children;

		if (border) {
			if (border.style) {
				attributes = [['style', border.style]];
			}
			if (border.color) {
				children = [formatUtils.exportColor(border.color)];
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

Borders.prototype.init = function () {
	this.formats.push(
		{format: this.canon({})}
	);
};

Borders.prototype.canon = Borders.canon;

Borders.prototype.merge = function (formatTo, formatFrom) {
	formatTo = formatTo || {};

	if (formatFrom) {
		_.forEach(BORDERS, function (name) {
			var borderFrom = formatFrom[name];

			if (borderFrom && (borderFrom.style || borderFrom.color)) {
				formatTo[name] = {
					style: borderFrom.style,
					color: borderFrom.color
				};
			} else if (!formatTo[name]) {
				formatTo[name] = {};
			}
		});
	}
	return formatTo;
};

Borders.prototype.exportFormat = Borders.exportFormat;

module.exports = Borders;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22,"./utils":25}],16:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var StylePart = require('./stylePart');
var alignment = require('./alignment');
var protection = require('./protection');
var util = require('../util');
var toXMLString = require('../XMLString');

var ALLOWED_PARTS = ['format', 'fill', 'border', 'font'];
var XLS_NAMES = ['numFmtId', 'fillId', 'borderId', 'fontId'];

function Cells(styles) {
	StylePart.call(this, styles, 'cellXfs', 'format');

	this.init();
	this.lastId = this.formats.length;
	this.exportEmpty = false;
}

util.inherits(Cells, StylePart);

Cells.prototype.init = function () {
	this.formats.push(
		{format: this.canon({})}
	);
};

Cells.prototype.canon = function (format, flags) {
	var result = {};
	var alignmentValue;
	var protectionValue;

	if (format.format) {
		result.format = this.styles.numberFormats.add(format.format);
	}
	if (format.font) {
		result.font = this.styles.fonts.add(format.font);
	}
	if (format.pattern) {
		result.fill = this.styles.fills.add(format.pattern, null, {fillType: 'pattern'});
	} else if (format.gradient) {
		result.fill = this.styles.fills.add(format.gradient, null, {fillType: 'gradient'});
	} else if (flags && flags.merge && format.fill) {
		result.fill = this.styles.fills.add(format.fill, null, flags);
	}
	if (format.border) {
		result.border = this.styles.borders.add(format.border);
	}
	alignmentValue = alignment.canon(format);
	if (alignmentValue) {
		result.alignment = alignmentValue;
	}
	protectionValue = protection.canon(format);
	if (protectionValue) {
		result.protection = protectionValue;
	}
	if (format.fillOut) {
		result.fillOut = format.fillOut;
	}
	return result;
};

Cells.prototype.fullGet = function (format) {
	var result = {};

	if (this.getId(format)) {
		format = this.get(format);
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
	} else {
		result = this.canon(format);
	}
	return result;
};

Cells.prototype.cutVisible = function (format) {
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
};

Cells.prototype.merge = function (formatTo, formatFrom) {
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
};

Cells.prototype.exportFormat = function (format) {
	var styles = this.styles;
	var attributes = [];
	var children = [];

	if (format.alignment) {
		children.push(alignment.exportFormat(format.alignment));
		attributes.push(['applyAlignment', 'true']);
	}
	if (format.protection) {
		children.push(protection.exportFormat(format.protection));
		attributes.push(['applyProtection', 'true']);
	}

	_.forEach(format, function (value, key) {
		var xlsName;

		if (_.includes(ALLOWED_PARTS, key)) {
			xlsName = XLS_NAMES[_.indexOf(ALLOWED_PARTS, key)];

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
};

module.exports = Cells;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./alignment":14,"./protection":21,"./stylePart":22}],17:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var StylePart = require('./stylePart');
var util = require('../util');
var toXMLString = require('../XMLString');

var PATTERN_TYPES = ['none', 'solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625',
	'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis',
	'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp',	'lightGrid', 'lightTrellis'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fills.aspx
function Fills(styles) {
	StylePart.call(this, styles, 'fills', 'fill');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Fills, StylePart);

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

Fills.exportFormat = function (format) {
	var children;

	if (format.fillType === 'pattern') {
		children = [exportPatternFill(format)];
	} else {
		children = [exportGradientFill(format)];
	}

	return toXMLString({
		name: 'fill',
		children: children
	});
};

function exportPatternFill(format) {
	var attributes = [
		['patternType', format.patternType]
	];
	var children = [
		toXMLString({
			name: 'fgColor',
			attributes: [
				['rgb', format.fgColor]
			]
		}),
		toXMLString({
			name: 'bgColor',
			attributes: [
				['rgb', format.bgColor]
			]
		})
	];

	return toXMLString({
		name: 'patternFill',
		attributes: attributes,
		children: children
	});
}

function exportGradientFill(format) {
	var attributes = [];
	var children = [];
	var attrs;

	if (format.degree) {
		attributes.push(['degree', format.degree]);
	} else if (format.left) {
		attributes.push(['type', 'path']);
		attributes.push(['left', format.left]);
		attributes.push(['right', format.right]);
		attributes.push(['top', format.top]);
		attributes.push(['bottom', format.bottom]);
	}

	attrs = [['rgb', format.start]];
	children.push(toXMLString({
		name: 'stop',
		attributes: [
			['position', 0]
		],
		children : [toXMLString({name: 'color',	attributes: attrs})]
	}));

	attrs = [['rgb', format.end]];
	children.push(toXMLString({
		name: 'stop',
		attributes: [
			['position', 1]
		],
		children : [toXMLString({name: 'color',	attributes: attrs})]
	}));

	return toXMLString({
		name: 'gradientFill',
		attributes: attributes,
		children: children
	});
}

Fills.prototype.init = function () {
	this.formats.push(
		{format: this.canon({type: 'none'}, {fillType: 'pattern'})},
		{format: this.canon({type: 'gray125'}, {fillType: 'pattern'})}
	);
};

Fills.prototype.canon = Fills.canon;

Fills.prototype.merge = function (formatTo, formatFrom) {
	return formatFrom || formatTo;
};

Fills.prototype.exportFormat = Fills.exportFormat;

module.exports = Fills;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22}],18:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var StylePart = require('./stylePart');
var util = require('../util');
var formatUtils = require('./utils');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fonts.aspx
function Fonts(styles) {
	StylePart.call(this, styles, 'fonts', 'font');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Fonts, StylePart);

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

Fonts.exportFormat = function (format) {
	var children = [];
	var attrs;

	if (format.size) {
		children.push(toXMLString({
			name: 'sz',
			attributes: [
				['val', format.size]
			]
		}));
	}
	if (format.fontName) {
		children.push(toXMLString({
			name: 'name',
			attributes: [
				['val', format.fontName]
			]
		}));
	}
	if (_.has(format, 'bold')) {
		if (format.bold) {
			children.push(toXMLString({name: 'b'}));
		} else {
			children.push(toXMLString({name: 'b', attributes: [['val', 0]]}));
		}
	}
	if (_.has(format, 'italic')) {
		if (format.italic) {
			children.push(toXMLString({name: 'i'}));
		} else {
			children.push(toXMLString({name: 'i', attributes: [['val', 0]]}));
		}
	}
	if (format.vertAlign) {
		children.push(toXMLString({
			name: 'vertAlign',
			attributes: [
				['val', format.vertAlign]
			]
		}));
	}
	if (format.underline) {
		attrs = null;

		if (format.underline !== true) {
			attrs = [
				['val', format.underline]
			];
		}
		children.push(toXMLString({
			name: 'u',
			attributes: attrs
		}));
	}
	if (_.has(format, 'strike')) {
		if (format.strike) {
			children.push(toXMLString({name: 'strike'}));
		} else {
			children.push(toXMLString({name: 'strike', attributes: [['val', 0]]}));
		}
	}
	if (format.shadow) {
		children.push(toXMLString({name: 'shadow'}));
	}
	if (format.outline) {
		children.push(toXMLString({name: 'outline'}));
	}
	if (format.color) {
		children.push(formatUtils.exportColor(format.color));
	}

	return toXMLString({
		name: 'font',
		children: children
	});
};

Fonts.prototype.init = function () {
	this.formats.push(
		{format: this.canon({})}
	);
};

Fonts.prototype.canon = Fonts.canon;

Fonts.prototype.merge = function (formatTo, formatFrom) {
	var result = _.assign(formatTo, formatFrom);

	result.color = formatFrom && formatFrom.color || formatTo && formatTo.color;
	return result;
};

Fonts.prototype.exportFormat = Fonts.exportFormat;

module.exports = Fonts;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22,"./utils":25}],19:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
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
}

Styles.prototype.addFormat = function (format, name) {
	return this.cells.add(format, name);
};

Styles.prototype._get = function (name) {
	return this.cells.get(name);
};

Styles.prototype._getId = function (name) {
	return this.cells.getId(name);
};

Styles.prototype._addInvisibleFormat = function (format) {
	var style = this.cells.cutVisible(this.cells.fullGet(format));

	if (!_.isEmpty(style)) {
		return this.addFormat(style);
	}
};

Styles.prototype._merge = function (columnFormat, rowFormat, cellFormat) {
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
		return this.cells.add(format, null, {merge: true});
	}
};

Styles.prototype.addFontFormat = function (format, name) {
	return this.fonts.add(format, name);
};

Styles.prototype.addBorderFormat = function (format, name) {
	return this.borders.add(format, name);
};

Styles.prototype.addPatternFormat = function (format, name) {
	return this.fills.add(format, name, {fillType: 'pattern'});
};

Styles.prototype.addGradientFormat = function (format, name) {
	return this.fills.add(format, name, {fillType: 'gradient'});
};

Styles.prototype.addNumberFormat = function (format, name) {
	return this.numberFormats.add(format, name);
};

Styles.prototype.addTableFormat = function (format, name) {
	return this.tables.add(format, name);
};

Styles.prototype.addTableElementFormat = function (format, name) {
	return this.tableElements.add(format, name);
};

Styles.prototype.setDefaultTableStyle = function (name) {
	this.tables.defaultTableStyle = name;
};

Styles.prototype.export = function () {
	return toXMLString({
		name: 'styleSheet',
		ns: 'spreadsheetml',
		children: [
			this.numberFormats.export(),
			this.fonts.export(),
			this.fills.export(),
			this.borders.export(),
			this.cells.export(),
			this.tableElements.export(),
			this.tables.export()
		]
	});
};

module.exports = Styles;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"./borders":15,"./cells":16,"./fills":17,"./fonts":18,"./numberFormats":20,"./tableElements":23,"./tables":24}],20:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var StylePart = require('./stylePart');
var util = require('../util');
var toXMLString = require('../XMLString');

var PREDEFINED = {
	date: 14, //mm-dd-yy
	time: 21  //h:mm:ss
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats.aspx
function NumberFormats(styles) {
	StylePart.call(this, styles, 'numFmts', 'numberFormat');

	this.init();
	this.lastId = 164;
}

util.inherits(NumberFormats, StylePart);

NumberFormats.canon = function (format) {
	return format;
};

NumberFormats.exportFormat = function (format, styleFormat) {
	var attributes = [
		['numFmtId', styleFormat.formatId],
		['formatCode', format]
	];

	return toXMLString({
		name: 'numFmt',
		attributes: attributes
	});
};

NumberFormats.prototype.init = function () {
	var self = this;

	_.forEach(PREDEFINED, function (formatId, format) {
		self.formatsByNames[format] = {
			formatId: formatId,
			format: format
		};
	});
};

NumberFormats.prototype.canon = NumberFormats.canon;

NumberFormats.prototype.merge = function (formatTo, formatFrom) {
	return formatFrom || formatTo;
};

NumberFormats.prototype.exportFormat = NumberFormats.exportFormat;

module.exports = NumberFormats;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22}],21:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
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

function exportProtection(format) {
	return toXMLString({
		name: 'protection',
		attributes: _.toPairs(format)
	});
}

module.exports = {
	canon: canon,
	merge: merge,
	exportFormat: exportProtection
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],22:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var toXMLString = require('../XMLString');

function StylePart(styles, exportName, formatName) {
	this.styles = styles;
	this.exportName = exportName;
	this.formatName = formatName;
	this.lastName = 1;
	this.lastId = 0;
	this.exportEmpty = true;
	this.formats = [];
	this.formatsByData = Object.create(null);
	this.formatsByNames = Object.create(null);
}

StylePart.prototype.add = function (format, name, flags) {
	var canonFormat;
	var stringFormat;
	var styleFormat;

	if (name && this.formatsByNames[name]) {
		canonFormat = this.canon(format, flags);
		stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;

		if (stringFormat !== this.formatsByNames[name].stringFormat) {
			this._add(canonFormat, stringFormat, name);
		}
		return name;
	}

	//first argument is format name
	if (!name && _.isString(format) && this.formatsByNames[format]) {
		return format;
	}

	canonFormat = this.canon(format, flags);
	stringFormat = _.isObject(canonFormat) ? JSON.stringify(canonFormat) : canonFormat;
	styleFormat = this.formatsByData[stringFormat];

	if (!styleFormat) {
		styleFormat = this._add(canonFormat, stringFormat, name);
	} else if (name && !this.formatsByNames[name]) {
		styleFormat.name = name;
		this.formatsByNames[name] = styleFormat;
	}
	return styleFormat.name;
};

StylePart.prototype._add = function (canonFormat, stringFormat, name) {
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
};

StylePart.prototype.canon = function (format) {
	return format;
};

StylePart.prototype.get = function (format) {
	if (_.isString(format)) {
		var styleFormat = this.formatsByNames[format];

		return styleFormat ? styleFormat.format : format;
	}
	return format;
};

StylePart.prototype.getId = function (name) {
	var styleFormat = this.formatsByNames[name];

	return styleFormat ? styleFormat.formatId : null;
};

StylePart.prototype.export = function () {
	if (this.exportEmpty !== false || this.formats.length) {
		var self = this;
		var attributes = [
			['count', this.formats.length]
		];
		var children = _.map(this.formats, function (format) {
			return self.exportFormat(format.format, format);
		});

		this.exportCollectionExt(attributes, children);

		return toXMLString({
			name: this.exportName,
			attributes: attributes,
			children: children
		});
	}
	return '';
};

StylePart.prototype.exportCollectionExt = function () {};

StylePart.prototype.exportFormat = function () {
	return '';
};

module.exports = StylePart;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],23:[function(require,module,exports){
'use strict';

var StylePart = require('./stylePart');
var util = require('../util');
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

util.inherits(TableElements, StylePart);

TableElements.prototype.canon = function (format) {
	var result = {};

	if (format.format) {
		result.format = NumberFormats.canon(format.format);
	}
	if (format.font) {
		result.font = Fonts.canon(format.font);
	}
	if (format.pattern) {
		result.fill = Fills.canon(format.pattern, {fillType: 'pattern', isTable: true});
	} else if (format.gradient) {
		result.fill = Fills.canon(format.gradient, {fillType: 'gradient'});
	}
	if (format.border) {
		result.border = Borders.canon(format.border);
	}
	result.alignment = alignment.canon(format);
	return result;
};

TableElements.prototype.exportFormat = function (format) {
	var children = [];

	if (format.font) {
		children.push(Fonts.exportFormat(format.font));
	}
	if (format.fill) {
		children.push(Fills.exportFormat(format.fill));
	}
	if (format.border) {
		children.push(Borders.exportFormat(format.border));
	}
	if (format.format) {
		children.push(NumberFormats.exportFormat(format.format));
	}
	if (format.alignment && format.alignment.length) {
		children.push(alignment.export(format.alignment));
	}

	return toXMLString({
		name: 'dxf',
		children: children
	});
};

module.exports = TableElements;

},{"../XMLString":2,"../util":27,"./alignment":14,"./borders":15,"./fills":17,"./fonts":18,"./numberFormats":20,"./stylePart":22}],24:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var StylePart = require('./stylePart');
var util = require('../util');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestylevalues.aspx
var ELEMENTS = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn',
	'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe',
	'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell'];
var SIZED_ELEMENTS = ['firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyles.aspx
function Tables(styles) {
	StylePart.call(this, styles, 'tableStyles', 'table');

	this.exportEmpty = false;
}

util.inherits(Tables, StylePart);

Tables.prototype.canon = function (format) {
	var result = {};
	var styles = this.styles;

	_.forEach(format, function (value, key) {
		if (_.includes(ELEMENTS, key)) {
			var style;
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
};

Tables.prototype.exportCollectionExt = function (attributes) {
	if (this.styles.defaultTableStyle) {
		attributes.push(['defaultTableStyle', this.styles.defaultTableStyle]);
	}
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyleelement.aspx
Tables.prototype.exportFormat = function (format, styleFormat) {
	var styles = this.styles;
	var attributes = [
		['name', styleFormat.name],
		['pivot', 0]
	];
	var children = [];

	_.forEach(format, function (value, key) {
		var style = value.style;
		var size = value.size;
		var attributes = [
			['type', key],
			['dxfId', styles.tableElements.getId(style)]
		];

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
};

module.exports = Tables;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22}],25:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var toXMLString = require('../XMLString');

function exportColor(color) {
	var attributes;

	if (_.isString(color)) {
		return toXMLString({
			name: 'color',
			attributes: [
				['rgb', color]
			]
		});
	} else {
		attributes = [];
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
	exportColor: exportColor
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],26:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('./util');
var toXMLString = require('./XMLString');

function Table(worksheet, config) {
	this.worksheet = worksheet;
	this.common = worksheet.common;

	var id = this.common.uniqueIdSeparated('Table');

	this.tableId = id.id;
	this.objectId = id.space + id.id;
	this.name = this.objectId;
	this.displayName = this.objectId;
	this.headerRowBorderDxfId = null;
	this.headerRowCount = 1;
	this.headerRowDxfId = null;
	this.beginCell = null;
	this.endCell = null;
	this.totalsRowCount = 0;
	this.totalRow = null;
	this.themeStyle = null;

	_.extend(this, config);
}

Table.prototype.end = function () {
	return this.worksheet;
};

Table.prototype.setReferenceRange = function (beginCell, endCell) {
	this.beginCell = util.canonCell(beginCell);
	this.endCell = util.canonCell(endCell);
	return this;
};

Table.prototype.addTotalRow = function (totalRow) {
	this.totalRow = totalRow;
	this.totalsRowCount = 1;
	return this;
};

Table.prototype.setTheme = function (theme) {
	this.themeStyle = theme;
	return this;
};

/**
 * Expects an object with the following properties:
 * caseSensitive (boolean)
 * dataRange
 * columnSort (assumes true)
 * sortDirection
 * sortRange (defaults to dataRange)
 */
Table.prototype.setSortState = function (state) {
	this.sortState = state;
	return this;
};

Table.prototype._prepare = function (worksheetData) {
	var SUB_TOTAL_FUNCTIONS = ['average', 'countNums', 'count', 'max', 'min', 'stdDev', 'sum', 'var'];
	var SUB_TOTAL_NUMS = [101, 102, 103, 104, 105, 107, 109, 110];

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
			var funcIndex;

			if (typeof headerValue === 'object') {
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
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.table.aspx
Table.prototype._export = function () {
	var attributes = [
		['id', this.tableId],
		['name', this.name],
		['displayName', this.displayName]
	];
	var children = [];
	var end = util.letterToPosition(this.endCell);
	var ref = this.beginCell + ':' + util.positionToLetter(end.x, end.y + this.totalsRowCount);

	attributes.push(['ref', ref]);
	attributes.push(['totalsRowCount', this.totalsRowCount]);
	attributes.push(['headerRowCount', this.headerRowCount]);

	children.push(exportAutoFilter(this.beginCell, this.endCell));
	children.push(exportTableColumns(this.totalRow));
	children.push(exportTableStyleInfo(this.common, this.themeStyle));

	return toXMLString({
		name: 'table',
		ns: 'spreadsheetml',
		attributes: attributes,
		children: children
	});
};

function exportAutoFilter(beginCell, endCell) {
	return toXMLString({
		name: 'autoFilter',
		attributes: ['ref', beginCell + ':' + endCell]
	});
}

function exportTableColumns(totalRow) {
	var attributes = [
		['count', totalRow.length]
	];
	var children = _.map(totalRow, function (cell, index) {
		var attributes = [
			['id', index + 1],
			['name', cell.name]
		];

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

function exportTableStyleInfo(common, themeStyle) {
	var attributes = [
		['name', themeStyle]
	];
	var isRowStripes = false;
	var isColumnStripes = false;
	var isFirstColumn = false;
	var isLastColumn = false;
	var format = common.styles.tables.get(themeStyle);

	if (format) {
		isRowStripes = format.firstRowStripe || format.secondRowStripe;
		isColumnStripes = format.firstColumnStripe || format.secondColumnStripe;
		isFirstColumn = format.firstColumn;
		isLastColumn = format.lastColumn;
	}
	attributes.push(
		['showRowStripes', isRowStripes ? '1' : '0'],
		['showColumnStripes', isColumnStripes ? '1' : '0'],
		['showFirstColumn', isFirstColumn ? '1' : '0'],
		['showLastColumn', isLastColumn ? '1' : '0']
	);

	return toXMLString({
		name: 'tableStyleInfo',
		attributes: attributes
	});
}

module.exports = Table;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./util":27}],27:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var LETTER_REFS = {};

function positionToLetter(x, y) {
	var result = LETTER_REFS[x];

	if (!result) {
		var string = '';
		var num = x;
		var index;

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
	var i;
	var len;
	var charCode;

	for (i = 0, len = cell.length; i < len; i++) {
		charCode = cell.charCodeAt(i);
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

var util = {
	inherits: function (ctor, superCtor) {
		var Obj = function () {};
		Obj.prototype = superCtor.prototype;
		ctor.prototype = new Obj();
	},

	pixelsToEMUs: function (pixels) {
		return Math.round(pixels * 914400 / 96);
	},

	canonCell: function (cell) {
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

module.exports = util;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{}],28:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('./util');
var Common = require('./common');
var Relations = require('./relations');
var Worksheet = require('./worksheet');
var toXMLString = require('./XMLString');

function Workbook() {
	this.common = new Common();
	this.styles = this.common.styles;
	this.relations = new Relations(this.common);

	this.objectId = this.common.uniqueId('Workbook');
}

Workbook.prototype.addWorksheet = function (config) {
	var worksheet;

	config = _.defaults(config, {
		name: this.common.getNewWorksheetDefaultName()
	});
	worksheet = new Worksheet(this, config);
	this.common.addWorksheet(worksheet);
	this.relations.addRelation(worksheet, 'worksheet');

	return worksheet;
};

Workbook.prototype.addFormat = function (format, name) {
	return this.styles.addFormat(format, name);
};

Workbook.prototype.addFontFormat = function (format, name) {
	return this.styles.addFontFormat(format, name);
};

Workbook.prototype.addBorderFormat = function (format, name) {
	return this.styles.addBorderFormat(format, name);
};

Workbook.prototype.addPatternFormat = function (format, name) {
	return this.styles.addPatternFormat(format, name);
};

Workbook.prototype.addGradientFormat = function (format, name) {
	return this.styles.addGradientFormat(format, name);
};

Workbook.prototype.addNumberFormat = function (format, name) {
	return this.styles.addNumberFormat(format, name);
};

Workbook.prototype.addTableFormat = function (format, name) {
	return this.styles.addTableFormat(format, name);
};

Workbook.prototype.addTableElementFormat = function (format, name) {
	return this.styles.addTableElementFormat(format, name);
};

Workbook.prototype.setDefaultTableStyle = function (name) {
	this.styles.setDefaultTableStyle(name);
	return this;
};

Workbook.prototype.addImage = function (data, type, name) {
	return this.common.addImage(data, type, name);
};

Workbook.prototype._generateFiles = function (zip, canStream) {
	prepareWorksheets(this.common);

	exportWorksheets(zip, canStream, this.common);
	exportTables(zip, this.common);
	exportImages(zip, this.common);
	exportDrawings(zip, this.common);
	exportStyles(zip, this.relations, this.styles);
	exportSharedStrings(zip, canStream, this.relations, this.common);
	zip.file('[Content_Types].xml', createContentTypes(this.common));
	zip.file('_rels/.rels', createWorkbookRelationship());
	zip.file('xl/workbook.xml', this._export());
	zip.file('xl/_rels/workbook.xml.rels', this.relations.export());
};

Workbook.prototype._export = function () {
	return toXMLString({
		name: 'workbook',
		ns: 'spreadsheetml',
		attributes: [
			['xmlns:r', util.schemas.relationships]
		],
		children: [
			bookViewsXML(this.common),
			sheetsXML(this.relations, this.common),
			definedNamesXML(this.common)
		]
	});
};

function bookViewsXML(common) {
	var activeTab = 0;
	var activeWorksheetId;

	if (common.activeWorksheet) {
		activeWorksheetId = common.activeWorksheet.objectId;

		activeTab = Math.max(activeTab, _.findIndex(common.worksheets, function (worksheet) {
			return worksheet.objectId === activeWorksheetId;
		}));
	}

	return toXMLString({
		name: 'bookViews',
		children: [
			toXMLString({
				name: 'workbookView',
				attributes: [
					['activeTab', activeTab]
				]
			})
		]
	});
}

function sheetsXML(relations, common) {
	var maxWorksheetNameLength = 31;
	var children = _.map(common.worksheets, function (worksheet, index) {
		// Microsoft Excel (2007, 2013) do not allow worksheet names longer than 31 characters
		// if the worksheet name is longer, Excel displays an 'Excel found unreadable content...' popup when opening the file
		if (worksheet.name.length > maxWorksheetNameLength) {
			throw 'Microsoft Excel requires work sheet names to be less than ' + (maxWorksheetNameLength + 1) +
			' characters long, work sheet name "' + worksheet.name + '" is ' + worksheet.name.length + ' characters long';
		}

		return toXMLString({
			name: 'sheet',
			attributes: [
				['name', worksheet.name],
				['sheetId', index + 1],
				['r:id', relations.getRelationshipId(worksheet)],
				['state', worksheet.getState()]
			]
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
		var children = [];

		_.forEach(common.worksheets, function (worksheet, index) {
			var entry = worksheet._printTitles;

			if (entry && (entry.topTo >= 0 || entry.leftTo >= 0)) {
				var value = '';
				var name = worksheet.name;

				if (entry.topTo >= 0) {
					value = name + '!$' + (entry.topFrom + 1) + ':$' + (entry.topTo + 1);
					if (entry.leftTo >= 0) {
						value += ',';
					}
				}
				if (entry.leftTo >= 0) {
					value += name + '!$' + util.positionToLetter(entry.leftFrom + 1) +
						':$' + util.positionToLetter(entry.leftTo + 1);
				}

				children.push(toXMLString({
					name: 'definedName',
					value: value,
					attributes: [
						['name', '_xlnm.Print_Titles'],
						['localSheetId', index]
					]
				}));
			}
		});

		return toXMLString({
			name: 'definedNames',
			children: children
		});
	}
	return '';
}

function prepareWorksheets(common) {
	_.forEach(common.worksheets, function (worksheet) {
		worksheet._prepare();
	});
}

function exportWorksheets(zip, canStream, common) {
	_.forEach(common.worksheets, function (worksheet) {
		zip.file(worksheet.path, worksheet._export(canStream));
		zip.file(worksheet.relationsPath, worksheet.relations.export());
	});
}

function exportTables(zip, common) {
	_.forEach(common.tables, function (table) {
		zip.file(table.path, table._export());
	});
}

function exportImages(zip, common) {
	_.forEach(common.getImages(), function (image) {
		zip.file(image.path, image.data, {base64: true, binary: true});
		image.data = null;
	});
	common.removeImages();
}

function exportDrawings(zip, common) {
	_.forEach(common.drawings, function (drawing) {
		zip.file(drawing.path, drawing.export());
		zip.file(drawing.relationsPath, drawing.relations.export());
	});
}

function exportStyles(zip, relations, styles) {
	relations.addRelation(styles, 'stylesheet');
	zip.file('xl/styles.xml', styles.export());
}

function exportSharedStrings(zip, canStream, relations, common) {
	if (!common.sharedStrings.isEmpty()) {
		relations.addRelation(common.sharedStrings, 'sharedStrings');
		zip.file('xl/sharedStrings.xml', common.sharedStrings.export(canStream));
	}
}

function createContentTypes(common) {
	var children = [];

	children.push(toXMLString({
		name: 'Default',
		attributes: [
			['Extension', 'rels'],
			['ContentType', 'application/vnd.openxmlformats-package.relationships+xml']
		]
	}));
	children.push(toXMLString({
		name: 'Default',
		attributes: [
			['Extension', 'xml'],
			['ContentType', 'application/xml']
		]
	}));
	children.push(toXMLString({
		name: 'Override',
		attributes: [
			['PartName', '/xl/workbook.xml'],
			['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml']
		]
	}));
	if (!common.sharedStrings.isEmpty()) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/sharedStrings.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml']
			]
		}));
	}
	children.push(toXMLString({
		name: 'Override',
		attributes: [
			['PartName', '/xl/styles.xml'],
			['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml']
		]
	}));

	_.forEach(common.worksheets, function (worksheet, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/worksheets/sheet' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']
			]
		}));
	});
	_.forEach(common.tables, function (table, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/tables/table' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']
			]
		}));
	});
	_.forEach(common.getExtensions(), function (contentType, extension) {
		children.push(toXMLString({
			name: 'Default',
			attributes: [
				['Extension', extension],
				['ContentType', contentType]
			]
		}));
	});
	_.forEach(common.drawings, function (drawing, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [
				['PartName', '/xl/drawings/drawing' + (index + 1) + '.xml'],
				['ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml']
			]
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
		children: [
			toXMLString({
				name: 'Relationship',
				attributes: [
					['Id', 'rId1'],
					['Type', util.schemas.officeDocument],
					['Target', 'xl/workbook.xml']
				]
			})
		]
	});
}

module.exports = Workbook;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./common":4,"./relations":13,"./util":27,"./worksheet":32}],29:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var Drawings = require('../drawings');
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._drawings = null;
	},
	methods: {
		setImage: function (image, config) {
			return this._setDrawing(image, config, 'anchor');
		},
		setImageOneCell: function (image, config) {
			return this._setDrawing(image, config, 'oneCell');
		},
		setImageAbsolute: function (image, config) {
			return this._setDrawing(image, config, 'absolute');
		},
		_setDrawing: function (image, config, anchorType) {
			var name;

			if (!this._drawings) {
				this._drawings = new Drawings(this.common);

				this.common.addDrawings(this._drawings);
				this.relations.addRelation(this._drawings, 'drawingRelationship');
			}

			if (_.isObject(image)) {
				name = this.common.addImage(image.data, image.type);
			} else {
				name = image;
			}

			this._drawings.addImage(name, config, anchorType);
			return this;
		},
		_insertDrawing: function (colIndex, rowIndex, image) {
			var config;

			if (typeof image === 'string' || image.data) {
				this._setDrawing(image, {c: colIndex + 1, r: rowIndex + 1}, 'anchor');
			} else {
				config = image.config || {};
				config.cell = {c: colIndex + 1, r: rowIndex + 1};

				this._setDrawing(image.image, config, 'anchor');
			}
		},
		_exportDrawing: function () {
			if (this._drawings) {
				return toXMLString({
					name: 'drawing',
					attributes: [
						['r:id', this.relations.getRelationshipId(this._drawings)]
					]
				});
			}
			return '';
		}
	}
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../drawings":10}],30:[function(require,module,exports){
(function (global){
'use strict';

var Readable = require('stream').Readable;
var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('../util');
var toXMLString = require('../XMLString');

module.exports = {
	_export: function (canStream) {
		if (canStream) {
			return new WorksheetStream({
				worksheet: this
			});
		} else {
			return exportBeforeRows(this) +
				exportData(this) +
				exportAfterRows(this);
		}
	}
};

function WorksheetStream(options) {
	Readable.call(this, options);

	this.worksheet = options.worksheet;
	this.status = 0;
}

util.inherits(WorksheetStream, Readable || {});

WorksheetStream.prototype._read = function (size) {
	var stop = false;
	var worksheet = this.worksheet;

	if (this.status === 0) {
		stop = !this.push(exportBeforeRows(worksheet));

		this.status = 1;
		this.index = 0;
		this.len = worksheet.preparedData.length;
	}

	if (this.status === 1) {
		var s = '';
		var data = worksheet.preparedData;
		var preparedRows = worksheet.preparedRows;

		while (this.index < this.len && !stop) {
			while (this.index < this.len && s.length < size) {
				s += exportRow(data[this.index], preparedRows[this.index], this.index);
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
		this.push(exportAfterRows(worksheet));
		this.push(null);
	}
};

function exportBeforeRows(worksheet) {
	return util.xmlPrefix + '<worksheet xmlns="' + util.schemas.spreadsheetml +
		'" xmlns:r="' + util.schemas.relationships + '" xmlns:mc="' + util.schemas.markupCompat + '">' +
		exportDimension(worksheet.maxX, worksheet.maxY) +
		worksheet._exportSheetView() +
		exportColumns(worksheet.preparedColumns) +
		'<sheetData>';
}

function exportAfterRows(worksheet) {
	return '</sheetData>' +
		// 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
		// with Microsoft Excel (2007, 2013)
		worksheet._exportMergeCells() +
		worksheet._exportHyperlinks() +
		worksheet._exportPrint() +
		worksheet._exportTables() +
		// the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
		// to issue with Microsoft Excel (2007, 2013)
		worksheet._exportDrawing() +
		'</worksheet>';
}

function exportData(worksheet) {
	var data = worksheet.preparedData;
	var preparedRows = worksheet.preparedRows;
	var children = '';

	for (var i = 0, len = data.length; i < len; i++) {
		children += exportRow(data[i], preparedRows[i], i);
		data[i] = null;
	}
	return children;
}

function exportRow(dataRow, row, rowIndex) {
	var rowLen;
	var rowChildren = [];
	var colIndex;
	var value;
	var attrs;

	if (dataRow) {
		rowLen = dataRow.length;
		rowChildren = new Array(rowLen);

		for (colIndex = 0; colIndex < rowLen; colIndex++) {
			value = dataRow[colIndex];

			if (!value) {
				continue;
			}

			attrs = ' r="' + util.positionToLetter(colIndex + 1, rowIndex + 1) + '"';
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

function exportDimension(maxX, maxY) {
	var attributes = [];

	if (maxX !== 0) {
		attributes.push(
			['ref', 'A1:' + util.positionToLetter(maxX, maxY)]
		);
	} else {
		attributes.push(
			['ref', 'A1']
		);
	}

	return toXMLString({
		name: 'dimension',
		attributes: attributes
	});
}

function exportColumns(columns) {
	if (columns.length) {
		var children = _.map(columns, function (column, index) {
			column = column || {};

			var attributes = [
				['min', column.min || index + 1],
				['max', column.max || index + 1]
			];

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

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"stream":1}],31:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('../util');
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._hyperlinks = [];
	},
	methods: {
		setHyperlink: function (hyperlink) {
			hyperlink.objectId = this.common.uniqueId('hyperlink');
			this.relations.addRelation({
				objectId: hyperlink.objectId,
				target: hyperlink.location,
				targetMode: hyperlink.targetMode || 'External'
			}, 'hyperlink');
			this._hyperlinks.push(hyperlink);
			return this;
		},
		_insertHyperlink: function (colIndex, rowIndex, hyperlink) {
			var location;
			var targetMode;
			var tooltip;

			if (typeof hyperlink === 'string') {
				location = hyperlink;
			} else {
				location = hyperlink.location;
				targetMode = hyperlink.targetMode;
				tooltip = hyperlink.tooltip;
			}
			this.setHyperlink({
				cell: {c: colIndex + 1, r: rowIndex + 1},
				location: location,
				targetMode: targetMode,
				tooltip: tooltip
			});
		},
		_exportHyperlinks: function () {
			var relations = this.relations;

			if (this._hyperlinks.length > 0) {
				var children = _.map(this._hyperlinks, function (hyperlink) {
					var attributes = [
						['ref', util.canonCell(hyperlink.cell)],
						['r:id', relations.getRelationshipId(hyperlink)]
					];

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
	}
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27}],32:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var sheetView = require('./sheetView');
var hyperlinks = require('./hyperlinks');
var mergedCells = require('./mergedCells');
var print = require('./print');
var drawing = require('./drawing');
var tables = require('./tables');
var prepareExport = require('./prepareExport');
var worksheetExport = require('./export');
var Relations = require('../relations');

function Worksheet(workbook, config) {
	config = config || {};

	this.workbook = workbook;
	this.common = workbook.common;

	sheetView.init.call(this, config);
	hyperlinks.init.call(this);
	mergedCells.init.call(this);
	drawing.init.call(this);
	tables.init.call(this);
	print.init.call(this);

	this.objectId = this.common.uniqueId('Worksheet');
	this.data = [];
	this.columns = [];
	this.rows = [];

	this.name = config.name;
	this.state = config.state || 'visible';
	this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
	this.relations = new Relations(this.common);
}

_.assign(Worksheet.prototype, prepareExport);
_.assign(Worksheet.prototype, worksheetExport);
_.assign(Worksheet.prototype, sheetView.methods);
_.assign(Worksheet.prototype, hyperlinks.methods);
_.assign(Worksheet.prototype, mergedCells.methods);
_.assign(Worksheet.prototype, drawing.methods);
_.assign(Worksheet.prototype, tables.methods);
_.assign(Worksheet.prototype, print.methods);

Worksheet.prototype.end = function () {
	return this.workbook;
};

Worksheet.prototype.setActive = function () {
	this.common.setActiveWorksheet(this);
	return this;
};

Worksheet.prototype.setVisible = function () {
	this.setState('visible');
	return this;
};

Worksheet.prototype.setHidden = function () {
	this.setState('hidden');
	return this;
};

/**
 * //http://www.datypic.com/sc/ooxml/t-ssml_ST_SheetState.html
 * @param state - visible | hidden | veryHidden
 */
Worksheet.prototype.setState = function (state) {
	this.state = state;
	return this;
};

Worksheet.prototype.getState = function () {
	return this.state;
};

Worksheet.prototype.setRows = function (startRow, rows) {
	var self = this;

	if (!rows) {
		rows = startRow;
		startRow = 0;
	} else {
		--startRow;
	}
	_.forEach(rows, function (row, i) {
		self.rows[startRow + i] = row;
	});
	return this;
};

Worksheet.prototype.setRow = function (rowIndex, meta) {
	this.rows[--rowIndex] = meta;
	return this;
};

/**
 * http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.aspx
 */
Worksheet.prototype.setColumns = function (startColumn, columns) {
	var self = this;

	if (!columns) {
		columns = startColumn;
		startColumn = 0;
	} else {
		--startColumn;
	}
	_.forEach(columns, function (column, i) {
		self.columns[startColumn + i] = column;
	});
	return this;
};

Worksheet.prototype.setColumn = function (columnIndex, column) {
	this.columns[--columnIndex] = column;
	return this;
};

Worksheet.prototype.setData = function (startRow, data) {
	var self = this;

	if (!data) {
		data = startRow;
		startRow = 0;
	} else {
		--startRow;
	}
	_.forEach(data, function (row, i) {
		self.data[startRow + i] = row;
	});
	return this;
};

module.exports = Worksheet;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../relations":13,"./drawing":29,"./export":30,"./hyperlinks":31,"./mergedCells":33,"./prepareExport":34,"./print":35,"./sheetView":36,"./tables":37}],33:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('../util');
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._mergedCells = [];
	},
	methods: {
		mergeCells: function (cell1, cell2) {
			this._mergedCells.push([cell1, cell2]);
			return this;
		},
		_insertMergeCells: function (dataRow, colIndex, rowIndex, colSpan, rowSpan) {
			var i, j;
			var row;

			if (colSpan) {
				for (j = 0; j < colSpan; j++) {
					dataRow.splice(colIndex + 1, 0, {style: null, type: 'empty'});
				}
			}
			if (rowSpan) {
				colSpan += 1;

				for (i = 0; i < rowSpan; i++) {
					//todo: original data changed
					row = this.data[rowIndex + i + 1];

					if (!row) {
						row = [];
						this.data[rowIndex + i + 1] = row;
					}

					if (row.length > colIndex) {
						for (j = 0; j < colSpan; j++) {
							row.splice(colIndex, 0, {style: null, type: 'empty'});
						}
					} else {
						for (j = 0; j < colSpan; j++) {
							row[colIndex + j] = {style: null, type: 'empty'};
						}
					}
				}
			}
		},
		_exportMergeCells: function () {
			if (this._mergedCells.length > 0) {
				var children = _.map(this._mergedCells, function (mergeCell) {
					return toXMLString({
						name: 'mergeCell',
						attributes: [
							['ref', util.canonCell(mergeCell[0]) + ':' + util.canonCell(mergeCell[1])]
						]
					});
				});

				return toXMLString({
					name: 'mergeCells',
					attributes: [
						['count', this._mergedCells.length]
					],
					children: children
				});
			}
			return '';
		}
	}
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27}],34:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);

module.exports = {
	_prepare: function () {
		var rowIndex;
		var len;
		var maxX = 0;
		var preparedDataRow;

		this.preparedData = [];
		this.preparedColumns = [];
		this.preparedRows = [];

		this._prepareTables();
		prepareColumns(this);
		prepareRows(this);

		for (rowIndex = 0, len = this.data.length; rowIndex < len; rowIndex++) {
			preparedDataRow = prepareDataRow(this, rowIndex);

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
};

function prepareColumns(worksheet) {
	var styles = worksheet.common.styles;

	_.forEach(worksheet.columns, function (column, index) {
		var preparedColumn;

		if (column) {
			preparedColumn = _.clone(column);

			if (column.style) {
				preparedColumn.style = styles.addFormat(column.style);
				if (styles._get(preparedColumn.style).fillOut) {
					preparedColumn.styleId = styles._getId(preparedColumn.style);
				} else {
					preparedColumn.styleId = styles._getId(styles._addInvisibleFormat(preparedColumn.style));
				}
			}
			worksheet.preparedColumns[index] = preparedColumn;
		}
	});
}

function prepareRows(worksheet) {
	var styles = worksheet.common.styles;

	_.forEach(worksheet.rows, function (row, index) {
		var preparedRow;

		if (row) {
			preparedRow = _.clone(row);

			if (row.style) {
				preparedRow.style = styles.addFormat(row.style);
			}
			worksheet.preparedRows[index] = preparedRow;
		}
	});
}

function prepareDataRow(worksheet, rowIndex) {
	var common = worksheet.common;
	var styles = common.styles;
	var row = worksheet.preparedRows[rowIndex];
	var dataRow = worksheet.data[rowIndex];
	var preparedDataRow = [];
	var rowStyle = null;
	var column;
	var colIndex;
	var value;
	var cellValue;
	var cellStyle;
	var cellType;
	var cellFormula;
	var isString;
	var date;

	if (dataRow) {
		if (!_.isArray(dataRow)) {
			row = mergeDataRowToRow(worksheet, row, dataRow);
			dataRow = dataRow.data;
		}
		if (row) {
			rowStyle = row.style || null;
		}

		for (colIndex = 0; colIndex < dataRow.length; colIndex++) {
			column = worksheet.preparedColumns[colIndex];
			value = dataRow[colIndex];

			cellStyle = null;
			cellFormula = null;
			isString = false;
			if (_.isDate(value)) {
				cellValue = value;
				cellType = 'date';
			} else if (value && typeof value === 'object') {
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
				if (column && column.type) {
					cellType = column.type;
				} else if (typeof cellValue === 'number') {
					cellType = 'number';
				} else if (typeof cellValue === 'string') {
					cellType = 'string';
				}
			}

			if (cellType === 'string') {
				cellValue = common.addString(cellValue);
				isString = true;
			} else if (cellType === 'date' || cellType === 'time') {
				date = 25569.0 + ((_.isDate(cellValue) ? cellValue.valueOf() : cellValue) - worksheet.timezoneOffset) /
					(60 * 60 * 24 * 1000);
				if (_.isFinite(date)) {
					cellValue = date;
				} else {
					cellValue = common.addString(String(cellValue));
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
		if (row.style) {
			if (styles._get(row.style).fillOut) {
				row.styleId = styles._getId(row.style);
			} else {
				row.styleId = styles._getId(styles._addInvisibleFormat(row.style));
			}
		}
		worksheet.preparedRows[rowIndex] = row;
	}

	return preparedDataRow;
}

function mergeDataRowToRow(worksheet, row, dataRow) {
	row = row || {};
	row.height = dataRow.height || row.height;
	row.style = dataRow.style ? worksheet.common.styles.addFormat(dataRow.style) : row.style;
	row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;

	return row;
}

function insertEmbedded(worksheet, dataRow, value, colIndex, rowIndex) {
	var colSpan;
	var rowSpan;

	if (value.hyperlink) {
		worksheet._insertHyperlink(colIndex, rowIndex, value.hyperlink);
	}
	if (value.image) {
		worksheet._insertDrawing(colIndex, rowIndex, value.image);
	}
	if (value.colspan || value.rowspan) {
		colSpan = (value.colspan || 1) - 1;
		rowSpan = (value.rowspan || 1) - 1;

		worksheet.mergeCells({c: colIndex + 1, r: rowIndex + 1},
			{c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan});
		worksheet._insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan);
	}
}

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{}],35:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._headers = [];
		this._footers = [];
	},
	methods: {
		/**
		 * Expects an array length of three.
		 * @param {Array} headers [left, center, right]
		 */
		setHeader: function (headers) {
			if (!_.isArray(headers)) {
				throw 'Invalid argument type - setHeader expects an array of three instructions';
			}
			this._headers = headers;
			return this;
		},
		/**
		 * Expects an array length of three.
		 * @param {Array} footers [left, center, right]
		 */
		setFooter: function (footers) {
			if (!_.isArray(footers)) {
				throw 'Invalid argument type - setFooter expects an array of three instructions';
			}
			this._footers = footers;
			return this;
		},
		/**
		 * Set page details in inches.
		 */
		setPageMargin: function (margin) {
			this._margin = _.defaults(margin, {
				left: 0.7,
				right: 0.7,
				top: 0.75,
				bottom: 0.75,
				header: 0.3,
				footer: 0.3
			});
			return this;
		},
		/**
		 * http://www.datypic.com/sc/ooxml/t-ssml_ST_Orientation.html
		 *
		 * Can be one of 'portrait' or 'landscape'.
		 *
		 * @param {String} orientation
		 */
		setPageOrientation: function (orientation) {
			this._orientation = orientation;
			return this;
		},
		/**
		 * Set rows to repeat for print
		 *
		 * @param {int|[int, int]} params - number of rows to repeat from the top | [first, last] repeat rows
		 */
		setPrintTitleTop: function (params) {
			this._printTitles = this._printTitles || {};

			if (_.isObject(params)) {
				this._printTitles.topFrom = params[0];
				this._printTitles.topTo = params[1];
			} else {
				this._printTitles.topFrom = 0;
				this._printTitles.topTo = params - 1;
			}
			return this;
		},
		/**
		 * Set columns to repeat for print
		 *
		 * @param {int|[int, int]} params - number of columns to repeat from the left | [first, last] repeat columns
		 */
		setPrintTitleLeft: function (params) {
			this._printTitles = this._printTitles || {};

			if (_.isObject(params)) {
				this._printTitles.leftFrom = params[0];
				this._printTitles.leftTo = params[1];
			} else {
				this._printTitles.leftFrom = 0;
				this._printTitles.leftTo = params - 1;
			}
			return this;
		},
		_exportPrint: function () {
			return exportPageMargins(this._margin) +
				exportPageSetup(this._orientation) +
				exportHeaderFooter(this._headers, this._footers);
		}
	}
};

function exportPageMargins(margin) {
	if (margin) {
		return toXMLString({
			name: 'pageMargins',
			attributes: [
				['top', margin.top],
				['bottom', margin.bottom],
				['left', margin.left],
				['right', margin.right],
				['header', margin.header],
				['footer', margin.footer]
			]
		});
	}
	return '';
}

function exportPageSetup(orientation) {
	if (orientation) {
		return toXMLString({
			name: 'pageSetup',
			attributes: [
				['orientation', orientation]
			]
		});
	}
	return '';
}

function exportHeaderFooter(headers, footers) {
	if (headers.length > 0 || footers.length > 0) {
		var children = [];

		if (headers.length > 0) {
			children.push(exportHeader(headers));
		}
		if (footers.length > 0) {
			children.push(exportFooter(footers));
		}

		return toXMLString({
			name: 'headerFooter',
			children: children
		});
	}
	return '';
}

function exportHeader(headers) {
	return toXMLString({
		name: 'oddHeader',
		value: compilePageDetailPackage(headers)
	});
}

function exportFooter(footers) {
	return toXMLString({
		name: 'oddFooter',
		value: compilePageDetailPackage(footers)
	});
}

function compilePageDetailPackage(data) {
	data = data || '';

	return [
		'&L', compilePageDetailPiece(data[0] || ''),
		'&C', compilePageDetailPiece(data[1] || ''),
		'&R', compilePageDetailPiece(data[2] || '')
	].join('');
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

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],36:[function(require,module,exports){
(function (global){
'use strict';

/**
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sheetview.aspx
 */
var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('../util');
var toXMLString = require('../XMLString');

module.exports = {
	init: function (config) {
		this._pane = null;
		this._attributes = {
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
				value: null//A1
			},
			view: {
				value: 'normal'//normal | pageBreakPreview | pageLayout
			},
			windowProtection: {
				value: null,
				bool: true
			},
			zoomScale: {
				value: null//10-400
			},
			zoomScaleNormal: {
				value: null//10-400
			},
			zoomScalePageLayoutView: {
				value: null//10-400
			},
			zoomScaleSheetLayoutView: {
				value: null//10-400
			}
		};

		var self = this;

		_.forEach(config, function (value, name) {
			if (name === 'freeze') {
				self.freeze(value.col, value.row, value.cell, value.activePane);
			} else if (name === 'split') {
				self.split(value.x, value.y, value.cell, value.activePane);
			} else if (self._attributes[name]) {
				self._attributes[name].value = value;
			}
		});
	},
	methods: {
		setAttribute: function (name, value) {
			if (this._attributes[name]) {
				this._attributes[name].value = value;
			}
			return this;
		},
		/**
		 * Add froze pane
		 * @param col - column number: 0, 1, 2 ...
		 * @param row - row number: 0, 1, 2 ...
		 * @param cell? - 'A1' | {c: 1, r: 1}
		 * @param activePane? - topLeft | topRight | bottomLeft | bottomRight
		 */
		freeze: function (col, row, cell, activePane) {
			this._pane = {
				state: 'frozen',
				xSplit: col,
				ySplit: row,
				topLeftCell: util.canonCell(cell) || util.positionToLetter(col + 1, row + 1),
				activePane: activePane || 'bottomRight'
			};
			return this;
		},
		/**
		 * Add split pane
		 * @param x - Horizontal position of the split, in points; 0 (zero) if none
		 * @param y - Vertical position of the split, in points; 0 (zero) if none
		 * @param cell? - 'A1' | {c: 1, r: 1}
		 * @param activePane? - topLeft | topRight | bottomLeft | bottomRight
		 */
		split: function (x, y, cell, activePane) {
			this._pane = {
				state: 'split',
				xSplit: x * 20,
				ySplit: y * 20,
				topLeftCell: util.canonCell(cell) || 'A1',
				activePane: activePane || 'bottomRight'
			};
			return this;
		},
		_exportSheetView: function () {
			var attributes = [
				['workbookViewId', 0]
			];

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
				children: [
					toXMLString({
						name: 'sheetView',
						attributes: attributes,
						children: [
							exportPane(this._pane)
						]
					})
				]
			});
		}
	}
};

function exportPane(pane) {
	if (pane) {
		return toXMLString({
			name: 'pane',
			attributes: [
				['state', pane.state],
				['xSplit', pane.xSplit],
				['ySplit', pane.ySplit],
				['topLeftCell', pane.topLeftCell],
				['activePane', pane.activePane]
			]
		});
	}
	return '';
}

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27}],37:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var Table = require('../table');
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._tables = [];
	},
	methods: {
		addTable: function (config) {
			var table = new Table(this, config);

			this.common.addTable(table);
			this.relations.addRelation(table, 'table');
			this._tables.push(table);

			return table;
		},
		_prepareTables: function () {
			var data = this.data;

			_.forEach(this._tables, function (table) {
				table._prepare(data);
			});
		},
		_exportTables: function () {
			var relations = this.relations;

			if (this._tables.length > 0) {
				var children = _.map(this._tables, function (table) {
					return toXMLString({
						name: 'tablePart',
						attributes: [
							['r:id', relations.getRelationshipId(table)]
						]
					});
				});

				return toXMLString({
					name: 'tableParts',
					attributes: [
						['count', this._tables.length]
					],
					children: children
				});
			}
			return '';
		}
	}
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../table":26}]},{},[12])(12)
});