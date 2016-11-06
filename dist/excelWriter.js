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

Anchor.prototype._exportWithContent = function (content) {
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
},{"../XMLString":2,"../util":27}],4:[function(require,module,exports){
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

AnchorAbsolute.prototype._exportWithContent = function (content) {
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

},{"../XMLString":2,"../util":27}],5:[function(require,module,exports){
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

AnchorOneCell.prototype._exportWithContent = function (content) {
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
},{"../XMLString":2,"../util":27}],6:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var Paths = require('./paths');
var SharedStrings = require('./sharedStrings');
var StyleSheet = require('./styleSheet');
var util = require('./util');

function Common() {
	this.paths = new Paths();

	this.sharedStrings = new SharedStrings();
	this.paths.add(this.sharedStrings, 'sharedStrings.xml');

	this.styleSheet = new StyleSheet();
	this.paths.add(this.styleSheet, 'styles.xml');

	this.worksheets = [];
	this.tables = [];
	this.images = {};
	this.imageByNames = {};
	this.drawings = [];
}

Common.prototype.addWorksheet = function (worksheet) {
	var index = this.worksheets.length + 1;
	var path = 'worksheets/sheet' + index + '.xml';
	var relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

	worksheet.path = 'xl/' + path;
	worksheet.relationsPath = relationsPath;
	this.worksheets.push(worksheet);
	this.paths.add(worksheet, path);
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
	this.paths.add(table, '/' + path);
};

Common.prototype.addImage = function (data, type, name) {
	var image = this.images[data];

	if (!image) {
		type = type || '';

		var contentType;
		var id = util.uniqueId('image');
		var path = 'xl/media/image' + id + '.' + type;

		name = name || '_jsExcelWriter' + id;
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
			objectId: _.uniqueId('Image'),
			data: data,
			name: name,
			contentType: contentType,
			extension: type,
			path: path
		};
		this.paths.add(image, '/' + path);
		this.images[data] = image;
		this.imageByNames[name] = image;
	} else if (name && !this.imageByNames[name]) {
		image.name = name;
		this.imageByNames[name] = image;
	}
	return image.name;
};

Common.prototype.getImage = function (name) {
	return this.imageByNames[name];
};

Common.prototype.addDrawings = function (drawings) {
	var index = this.drawings.length + 1;
	var path = 'xl/drawings/drawing' + index + '.xml';
	var relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

	drawings.path = path;
	drawings.relationsPath = relationsPath;
	this.drawings.push(drawings);
	this.paths.add(drawings, '/' + path);
};

module.exports = Common;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./paths":9,"./sharedStrings":12,"./styleSheet":14,"./util":27}],7:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('./util');
var RelationshipManager = require('./relationshipManager');
var Picture = require('./picture');
var toXMLString = require('./XMLString');

function Drawings(common) {
	this.objectId = _.uniqueId('Drawings');
	this.drawings = [];
	this.relations = new RelationshipManager(common);
	this.common = common;
}

Drawings.prototype.addImage = function (name, config, anchorType) {
	var image = this.common.getImage(name);
	var imageRelationId = this.relations.addRelation(image, 'image');
	var picture = new Picture({
		image: image,
		imageRelationId: imageRelationId,
		config: config,
		anchorType: anchorType
	});

	this.drawings.push(picture);
};

Drawings.prototype._export = function () {
	var attributes = [
		['xmlns:a', util.schemas.drawing],
		['xmlns:r', util.schemas.relationships],
		['xmlns:xdr', util.schemas.spreadsheetDrawing]
	];
	var children = _.map(this.drawings, function (picture) {
		return picture._export();
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
},{"./XMLString":2,"./picture":10,"./relationshipManager":11,"./util":27}],8:[function(require,module,exports){
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
},{"./workbook":28}],9:[function(require,module,exports){
'use strict';

function Paths() {
	this.paths = {};
}

Paths.prototype.add = function (object, path) {
	this.paths[object.objectId] = path;
};

Paths.prototype.get = function (object) {
	return this.paths[object.objectId];
};

module.exports = Paths;

},{}],10:[function(require,module,exports){
'use strict';

var Anchor = require('./anchor/anchor');
var AnchorOneCell = require('./anchor/anchorOneCell');
var AnchorAbsolute = require('./anchor/anchorAbsolute');
var util = require('./util');
var toXMLString = require('./XMLString');

function Picture(config) {
	this.objectId = util.uniqueId('Picture');
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

Picture.prototype._export = function () {
	var picture = toXMLString({
		name: 'xdr:pic',
		children: [
			toXMLString({
				name: 'xdr:nvPicPr',
				children: [
					toXMLString({
						name: 'xdr:cNvPr',
						attributes: [
							['id', this.objectId],
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

	return this.anchor._exportWithContent(picture);
};

module.exports = Picture;

},{"./XMLString":2,"./anchor/anchor":3,"./anchor/anchorAbsolute":4,"./anchor/anchorOneCell":5,"./util":27}],11:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('./util');
var toXMLString = require('./XMLString');

function RelationshipManager(common) {
	this.relations = {};
	this.lastId = 1;
	this.paths = common.paths;
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

RelationshipManager.prototype._export = function () {
	var paths = this.paths;
	var children = _.map(this.relations, function (relation) {
		var attributes = [
			['Id', relation.relationId],
			['Type', relation.schema],
			['Target', relation.object.target || paths.get(relation.object)]
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
},{"./XMLString":2,"./util":27}],12:[function(require,module,exports){
(function (global){
'use strict';

var Readable = require('stream').Readable;
var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('./util');

function SharedStrings() {
	this.objectId = _.uniqueId('SharedStrings');
	this._strings = {};
	this._stringArray = [];
}

/**
 * Adds a string to the shared string file, and returns the ID of the
 * string which can be used to reference it in worksheets.
 *
 * @param string {String}
 * @return int
 */
SharedStrings.prototype.addString = function (string) {
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

SharedStrings.prototype._export = function (canStream) {
	if (canStream) {
		return new SharedStringsStream({
			strings: this._stringArray
		});
	} else {
		this._strings = null;

		var len = this._stringArray.length;
		var children = _.map(this._stringArray, function (string) {
			return '<si><t>' + _.escape(string) + '</t></si>';
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
		this._strings = null;
		stop = !this.push(getXMLBegin(this.strings.length));

		this.status = 1;
		this.index = 0;
		this.len = this.strings.length;
	}

	if (this.status === 1) {
		var s = '';
		while (this.index < this.len && !stop) {
			while (this.index < this.len && s.length < size) {
				s += '<si><t>' + _.escape(this.strings[this.index]) + '</t></si>';
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
},{"./util":27,"stream":1}],13:[function(require,module,exports){
(function (global){
'use strict';

/**
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sheetview.aspx
 */
var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('./util');
var toXMLString = require('./XMLString');

function SheetView(config) {
	config = config || {};
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
		} else if (self.attributes[name]) {
			self.attributes[name].value = value;
		}
	});
}

SheetView.prototype.setAttribute = function (name, value) {
	if (this.attributes[name]) {
		this.attributes[name].value = value;
	}
};

/**
 * Add froze pane
 * @param col - column number: 0, 1, 2 ...
 * @param row - row number: 0, 1, 2 ...
 * @param cell? - 'A1' | {c: 1, r: 1}
 * @param activePane? - topLeft | topRight | bottomLeft | bottomRight
 */
SheetView.prototype.freeze = function (col, row, cell, activePane) {
	this.pane = {
		state: 'frozen',
		xSplit: col,
		ySplit: row,
		topLeftCell: util.canonCell(cell) || util.positionToLetter(col + 1, row + 1),
		activePane: activePane || 'bottomRight'
	};
};

/**
 * Add split pane
 * @param x - Horizontal position of the split, in points; 0 (zero) if none
 * @param y - Vertical position of the split, in points; 0 (zero) if none
 * @param cell? - 'A1' | {c: 1, r: 1}
 * @param activePane? - topLeft | topRight | bottomLeft | bottomRight
 */
SheetView.prototype.split = function (x, y, cell, activePane) {
	this.pane = {
		state: 'split',
		xSplit: x * 20,
		ySplit: y * 20,
		topLeftCell: util.canonCell(cell) || 'A1',
		activePane: activePane || 'bottomRight'
	};
};

SheetView.prototype._export = function () {
	var attributes = [
		['workbookViewId', 0]
	];

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
		children: [
			toXMLString({
				name: 'sheetView',
				attributes: attributes,
				children: [
					exportPane(this.pane)
				]
			})
		]
	});
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

module.exports = SheetView;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./util":27}],14:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var numberFormats = require('./style/numberFormats');
var fonts = require('./style/fonts');
var fills = require('./style/fills');
var borders = require('./style/borders');
var cells = require('./style/cells');
var tables = require('./style/tables');
var tableElements = require('./style/tableElements');
var toXMLString = require('./XMLString');

function StyleSheet() {
	this.objectId = _.uniqueId('StyleSheet');
	this.numberFormats = new numberFormats.NumberFormats(this);
	this.fonts = new fonts.Fonts(this);
	this.fills = new fills.Fills(this);
	this.borders = new borders.Borders(this);
	this.cells = new cells.Cells(this);
	this.tableElements = new tableElements.TableElements(this);
	this.tables = new tables.Tables(this);
	this.defaultTableStyle = '';
}

StyleSheet.prototype.addFormat = function (format, name) {
	return this.cells.add(format, null, name);
};

StyleSheet.prototype.addFontFormat = function (format, name) {
	return this.fonts.add(format, null, name);
};

StyleSheet.prototype.addBorderFormat = function (format, name) {
	return this.borders.add(format, null, name);
};

StyleSheet.prototype.addPatternFormat = function (format, name) {
	return this.fills.add(format, 'pattern', name);
};

StyleSheet.prototype.addGradientFormat = function (format, name) {
	return this.fills.add(format, 'gradient', name);
};

StyleSheet.prototype.addNumberFormat = function (format, name) {
	return this.numberFormats.add(format, null, name);
};

StyleSheet.prototype.addTableFormat = function (format, name) {
	return this.tables.add(format, null, name);
};

StyleSheet.prototype.addTableElementFormat = function (format, name) {
	return this.tableElements.add(format, null, name);
};

StyleSheet.prototype.setDefaultTableStyle = function (name) {
	this.tables.defaultTableStyle = name;
};

StyleSheet.prototype._export = function () {
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

module.exports = StyleSheet;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./style/borders":16,"./style/cells":17,"./style/fills":18,"./style/fonts":19,"./style/numberFormats":20,"./style/tableElements":23,"./style/tables":24}],15:[function(require,module,exports){
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
	var result = [];

	if (format.horizontal && _.includes(HORIZONTAL, format.horizontal)) {
		result.push(['horizontal', format.horizontal]);
	}
	if (format.vertical && _.includes(VERTICAL, format.vertical)) {
		result.push(['vertical', format.vertical]);
	}
	if (format.indent) {
		result.push(['indent', format.indent]);
	}
	if (format.justifyLastLine) {
		result.push(['justifyLastLine', 1]);
	}
	if (_.has(format, 'readingOrder') && _.includes([0, 1, 2], format.readingOrder)) {
		result.push(['readingOrder', format.readingOrder]);
	}
	if (format.relativeIndent) {
		result.push(['relativeIndent', format.relativeIndent]);
	}
	if (format.shrinkToFit) {
		result.push(['shrinkToFit', 1]);
	}
	if (format.textRotation) {
		result.push(['textRotation', format.textRotation]);
	}
	if (format.wrapText) {
		result.push(['wrapText', 1]);
	}

	return result.length ? result : undefined;
}

function exportAlignment(attributes) {
	return toXMLString({
		name: 'alignment',
		attributes: attributes
	});
}

module.exports = {
	canon: canon,
	export: exportAlignment
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],16:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var StylePart = require('./stylePart');
var util = require('../util');
var formatUtils = require('./utils');
var toXMLString = require('../XMLString');

var BORDERS = ['left', 'right', 'top', 'bottom', 'diagonal'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borders.aspx
function Borders(styleSheet) {
	StylePart.call(this, styleSheet, 'borders', 'border');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Borders, StylePart);

Borders.prototype.init = function () {
	this.formats.push(
		{format: this.canon({})}
	);
};

Borders.prototype.canon = canon;
Borders.prototype.exportFormat = exportFormat;

function canon(format) {
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
}

function exportFormat(format) {
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
}

module.exports = {
	Borders: Borders,
	canon: canon,
	exportFormat: exportFormat
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22,"./utils":25}],17:[function(require,module,exports){
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

function Cells(styleSheet) {
	StylePart.call(this, styleSheet, 'cellXfs', 'format');

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

Cells.prototype.canon = function (format) {
	var result = {};

	if (format.format) {
		result.format = this.styleSheet.numberFormats.add(format.format);
	}
	if (format.font) {
		result.font = this.styleSheet.fonts.add(format.font);
	}
	if (format.pattern) {
		result.fill = this.styleSheet.fills.add(format.pattern, 'pattern');
	} else if (format.gradient) {
		result.fill = this.styleSheet.fills.add(format.gradient, 'gradient');
	}
	if (format.border) {
		result.border = this.styleSheet.borders.add(format.border);
	}
	result.alignment = alignment.canon(format);
	result.protection = protection.canon(format);
	return result;
};

Cells.prototype.exportFormat = function (format) {
	var styleSheet = this.styleSheet;
	var attributes = [];
	var children = [];

	if (format.alignment && format.alignment.length) {
		children.push(alignment.export(format.alignment));
		attributes.push(['applyAlignment', 'true']);
	}
	if (format.protection && format.protection.length) {
		children.push(protection.export(format.protection));
		attributes.push(['applyProtection', 'true']);
	}

	_.forEach(format, function (value, key) {
		var xlsName;

		if (_.includes(ALLOWED_PARTS, key)) {
			xlsName = XLS_NAMES[_.indexOf(ALLOWED_PARTS, key)];

			if (key === 'format') {
				attributes.push([xlsName, styleSheet.numberFormats.getId(value)]);
				attributes.push(['applyNumberFormat', 'true']);
			} else if (key === 'fill') {
				attributes.push([xlsName, styleSheet.fills.getId(value)]);
				attributes.push(['applyFill', 'true']);
			} else if (key === 'border') {
				attributes.push([xlsName, styleSheet.borders.getId(value)]);
				attributes.push(['applyBorder', 'true']);
			} else if (key === 'font') {
				attributes.push([xlsName, styleSheet.fonts.getId(value)]);
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

module.exports = {
	Cells: Cells
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./alignment":15,"./protection":21,"./stylePart":22}],18:[function(require,module,exports){
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
function Fills(styleSheet) {
	StylePart.call(this, styleSheet, 'fills', 'fill');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Fills, StylePart);

Fills.prototype.init = function () {
	this.formats.push({
		format: this.canon({type: 'none'}, 'pattern')
	}, {
		format: this.canon({type: 'gray125'}, 'pattern')
	});
};

Fills.prototype.canon = canon;
Fills.prototype.exportFormat = exportFormat;

function canon(format, type, isTable) {
	var result = {
		type: type
	};

	if (type === 'pattern') {
		var fgColor = format.color || 'FFFFFFFF';
		var bgColor = format.backColor || 'FFFFFFFF';

		result.patternType = _.includes(PATTERN_TYPES, format.type) ? format.type : 'solid';
		if (isTable && result.patternType === 'solid') {
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

function exportFormat(format) {
	var children;

	if (format.type === 'pattern') {
		children = [exportPatternFill(format)];
	} else {
		children = [exportGradientFill(format)];
	}

	return toXMLString({
		name: 'fill',
		children: children
	});
}

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

module.exports = {
	Fills: Fills,
	canon: canon,
	exportFormat: exportFormat
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22}],19:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var StylePart = require('./stylePart');
var util = require('../util');
var formatUtils = require('./utils');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fonts.aspx
function Fonts(styleSheet) {
	StylePart.call(this, styleSheet, 'fonts', 'font');

	this.init();
	this.lastId = this.formats.length;
}

util.inherits(Fonts, StylePart);

Fonts.prototype.init = function () {
	this.formats.push(
		{format: canon({})}
	);
};

Fonts.prototype.canon = canon;
Fonts.prototype.exportFormat = exportFormat;

function canon(format) {
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

function exportFormat(format) {
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
}

module.exports = {
	Fonts: Fonts,
	canon: canon,
	exportFormat: exportFormat
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22,"./utils":25}],20:[function(require,module,exports){
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
function NumberFormats(styleSheet) {
	StylePart.call(this, styleSheet, 'numFmts', 'numberFormat');

	this.init();
	this.lastId = 164;
}

util.inherits(NumberFormats, StylePart);

NumberFormats.prototype.init = function () {
	var self = this;

	_.forEach(PREDEFINED, function (formatId, format) {
		self.formatsByNames[format] = {formatId: formatId};
	});
};

NumberFormats.prototype.canon = canon;
NumberFormats.prototype.exportFormat = exportFormat;

function canon(format) {
	return format;
}

function exportFormat(format, styleFormat) {
	var attributes = [
		['numFmtId', styleFormat.formatId],
		['formatCode', format]
	];

	return toXMLString({
		name: 'numFmt',
		attributes: attributes
	});
}

module.exports = {
	NumberFormats: NumberFormats,
	canon: canon,
	exportFormat: exportFormat
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2,"../util":27,"./stylePart":22}],21:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.protection.aspx

function canon(format) {
	var result = [];

	if (_.has(format, 'locked') && !format.locked) {
		result.push(['locked', 0]);
	}
	if (format.hidden) {
		result.push(['hidden', 1]);
	}

	return result.length ? result : undefined;
}

function exportProtection(attributes) {
	return toXMLString({
		name: 'protection',
		attributes: attributes
	});
}

module.exports = {
	canon: canon,
	export: exportProtection
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":2}],22:[function(require,module,exports){
(function (global){
'use strict';

var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var toXMLString = require('../XMLString');

function StylePart(styleSheet, exportName, formatName) {
	this.styleSheet = styleSheet;
	this.exportName = exportName;
	this.formatName = formatName;
	this.lastName = 1;
	this.lastId = 0;
	this.exportEmpty = true;
	this.formats = [];
	this.formatsByData = {};
	this.formatsByNames = {};
}

StylePart.prototype.add = function (format, type, name) {
	var canonFormat;
	var stringFormat;
	var styleFormat;

	if (name && this.formatsByNames[name]) {
		canonFormat = this.canon(format, type);
		if (_.isObject(canonFormat)) {
			stringFormat = JSON.stringify(canonFormat);
		} else {
			stringFormat = canonFormat;
		}

		if (stringFormat !== this.formatsByNames[name].format) {
			this._add(canonFormat, stringFormat, name);
		}
		return name;
	}

	//first argument is format name
	if (!name && _.isString(format) && this.formatsByNames[format]) {
		return format;
	}

	canonFormat = this.canon(format, type);
	if (_.isObject(canonFormat)) {
		stringFormat = JSON.stringify(canonFormat);
	} else {
		stringFormat = canonFormat;
	}
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
		format: canonFormat
	};

	this.formats.push(styleFormat);
	this.formatsByData[stringFormat] = styleFormat;
	this.formatsByNames[name] = styleFormat;

	return styleFormat;
};

StylePart.prototype.get = function (name) {
	var styleFormat = this.formatsByNames[name];

	return styleFormat ? styleFormat.format : null;
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

StylePart.prototype.canon = function (format) {
	return format;
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
var numberFormats = require('./numberFormats');
var fonts = require('./fonts');
var fills = require('./fills');
var borders = require('./borders');
var alignment = require('./alignment');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.differentialformats.aspx
function TableElements(styleSheet) {
	StylePart.call(this, styleSheet, 'dxfs', 'tableElement');
}

util.inherits(TableElements, StylePart);

TableElements.prototype.canon = function (format) {
	var result = {};

	if (format.format) {
		result.format = numberFormats.canon(format.format);
	}
	if (format.font) {
		result.font = fonts.canon(format.font);
	}
	if (format.pattern) {
		result.fill = fills.canon(format.pattern, 'pattern', 'table');
	} else if (format.gradient) {
		result.fill = fills.canon(format.gradient, 'gradient');
	}
	if (format.border) {
		result.border = borders.canon(format.border);
	}
	result.alignment = alignment.canon(format);
	return result;
};

TableElements.prototype.exportFormat = function (format) {
	var children = [];

	if (format.font) {
		children.push(fonts.exportFormat(format.font));
	}
	if (format.fill) {
		children.push(fills.exportFormat(format.fill));
	}
	if (format.border) {
		children.push(borders.exportFormat(format.border));
	}
	if (format.format) {
		children.push(numberFormats.exportFormat(format.format));
	}
	if (format.alignment && format.alignment.length) {
		children.push(alignment.export(format.alignment));
	}

	return toXMLString({
		name: 'dxf',
		children: children
	});
};

module.exports = {
	TableElements: TableElements
};

},{"../XMLString":2,"../util":27,"./alignment":15,"./borders":16,"./fills":18,"./fonts":19,"./numberFormats":20,"./stylePart":22}],24:[function(require,module,exports){
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
function Tables(styleSheet) {
	StylePart.call(this, styleSheet, 'tableStyles', 'table');

	this.exportEmpty = false;
}

util.inherits(Tables, StylePart);

Tables.prototype.canon = function (format) {
	var result = {};
	var styleSheet = this.styleSheet;

	_.forEach(format, function (value, key) {
		if (_.includes(ELEMENTS, key)) {
			var style;
			var size = null;

			if (value.style) {
				style = styleSheet.tableElements.add(value.style);
				if (value.size > 1 && _.includes(SIZED_ELEMENTS, key)) {
					size = value.size;
				}
			} else {
				style = styleSheet.tableElements.add(value);
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
	if (this.styleSheet.defaultTableStyle) {
		attributes.push(['defaultTableStyle', this.styleSheet.defaultTableStyle]);
	}
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyleelement.aspx
Tables.prototype.exportFormat = function (format, styleFormat) {
	var styleSheet = this.styleSheet;
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
			['dxfId', styleSheet.tableElements.getId(style)]
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

module.exports = {
	Tables: Tables
};

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
	this.tableId = util.uniqueId('Table');
	this.objectId = 'Table' + this.tableId;
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
	var format = common.styleSheet.tables.get(themeStyle);

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
var idSpaces = {};
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
	/**
	 * Returns a number based on a namespace. So, running with 'Picture' will return 1. Run again, you will get 2. Run with 'Foo', you'll get 1.
	 * @param {String} space
	 * @returns {Number}
	 */
	uniqueId: function (space) {
		if (!idSpaces[space]) {
			idSpaces[space] = 1;
		}
		return idSpaces[space]++;
	},

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
var RelationshipManager = require('./relationshipManager');
var Worksheet = require('./worksheet');
var toXMLString = require('./XMLString');

function Workbook() {
	this.objectId = _.uniqueId('Workbook');

	this.common = new Common();
	this.relations = new RelationshipManager(this.common);
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
	return this.common.styleSheet.addFormat(format, name);
};

Workbook.prototype.addFontFormat = function (format, name) {
	return this.common.styleSheet.addFontFormat(format, name);
};

Workbook.prototype.addBorderFormat = function (format, name) {
	return this.common.styleSheet.addBorderFormat(format, name);
};

Workbook.prototype.addPatternFormat = function (format, name) {
	return this.common.styleSheet.addPatternFormat(format, name);
};

Workbook.prototype.addGradientFormat = function (format, name) {
	return this.common.styleSheet.addGradientFormat(format, name);
};

Workbook.prototype.addNumberFormat = function (format, name) {
	return this.common.styleSheet.addNumberFormat(format, name);
};

Workbook.prototype.addTableFormat = function (format, name) {
	return this.common.styleSheet.addTableFormat(format, name);
};

Workbook.prototype.addTableElementFormat = function (format, name) {
	return this.common.styleSheet.addTableElementFormat(format, name);
};

Workbook.prototype.setDefaultTableStyle = function (name) {
	this.common.styleSheet.setDefaultTableStyle(name);
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
	exportStyles(zip, this.relations, this.common);
	exportSharedStrings(zip, canStream, this.relations, this.common);
	zip.file('[Content_Types].xml', createContentTypes(this.common));
	zip.file('_rels/.rels', createWorkbookRelationship());
	zip.file('xl/workbook.xml', this._export());
	zip.file('xl/_rels/workbook.xml.rels', this.relations._export());
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
				['state', worksheet.state]
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
		zip.file(worksheet.relationsPath, worksheet.relations._export());
	});
}

function exportTables(zip, common) {
	_.forEach(common.tables, function (table) {
		zip.file(table.path, table._export());
	});
}

function exportImages(zip, common) {
	_.forEach(common.images, function (image) {
		zip.file(image.path, image.data, {base64: true, binary: true});
		image.data = null;
	});
	common.images = null;
}

function exportDrawings(zip, common) {
	_.forEach(common.drawings, function (drawing) {
		zip.file(drawing.path, drawing._export());
		zip.file(drawing.relationsPath, drawing.relations._export());
	});
}

function exportStyles(zip, relations, common) {
	relations.addRelation(common.styleSheet, 'stylesheet');
	zip.file('xl/styles.xml', common.styleSheet._export());
}

function exportSharedStrings(zip, canStream, relations, common) {
	if (!common.sharedStrings.isEmpty()) {
		relations.addRelation(common.sharedStrings, 'sharedStrings');
		zip.file('xl/sharedStrings.xml', common.sharedStrings._export(canStream));
	}
}

function createContentTypes(common) {
	var children = [];
	var extensions = {};

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
	_.forEach(common.imageByNames, function (image) {
		extensions[image.extension] = image.contentType;
	});
	_.forEach(extensions, function (contentType, extension) {
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
},{"./XMLString":2,"./common":6,"./relationshipManager":11,"./util":27,"./worksheet":29}],29:[function(require,module,exports){
(function (global){
'use strict';

var Readable = require('stream').Readable;
var _ = (typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null);
var util = require('./util');
var SheetView = require('./sheetView');
var RelationshipManager = require('./relationshipManager');
var Table = require('./table');
var Drawings = require('./drawings');
var toXMLString = require('./XMLString');

function Worksheet(workbook, config) {
	config = config || {};

	this.objectId = _.uniqueId('Worksheet');
	this.workbook = workbook;
	this.common = workbook.common;

	this.data = [];
	this.columns = [];
	this.rows = [];
	this.mergedCells = [];
	this.headers = [];
	this.footers = [];
	this.tables = [];
	this.drawings = null;
	this.hyperlinks = [];

	this.name = config.name;
	this.state = config.state || 'visible';
	this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;

	this.relations = new RelationshipManager(this.common);
	this.sheetView = new SheetView(config);
}

Worksheet.prototype.end = function () {
	return this.workbook;
};

Worksheet.prototype.addTable = function (config) {
	var table = new Table(this, config);

	this.common.addTable(table);
	this.relations.addRelation(table, 'table');
	this.tables.push(table);

	return table;
};

Worksheet.prototype.setImage = function (image, config) {
	this._setImage(image, config, 'anchor');
	return this;
};

Worksheet.prototype.setImageOneCell = function (image, config) {
	this._setImage(image, config, 'oneCell');
	return this;
};

Worksheet.prototype.setImageAbsolute = function (image, config) {
	this._setImage(image, config, 'absolute');
	return this;
};

Worksheet.prototype._setImage = function (image, config, anchorType) {
	var name;

	if (!this.drawings) {
		this.drawings = new Drawings(this.common);

		this.common.addDrawings(this.drawings);
		this.relations.addRelation(this.drawings, 'drawingRelationship');
	}

	if (_.isObject(image)) {
		name = this.common.addImage(image.data, image.type);
	} else {
		name = image;
	}

	this.drawings.addImage(name, config, anchorType);
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

/**
 * Expects an array length of three.
 * @param {Array} headers [left, center, right]
 */
Worksheet.prototype.setHeader = function (headers) {
	if (!_.isArray(headers)) {
		throw 'Invalid argument type - setHeader expects an array of three instructions';
	}
	this.headers = headers;
	return this;
};

/**
 * Expects an array length of three.
 * @param {Array} footers [left, center, right]
 */
Worksheet.prototype.setFooter = function (footers) {
	if (!_.isArray(footers)) {
		throw 'Invalid argument type - setFooter expects an array of three instructions';
	}
	this.footers = footers;
	return this;
};

/**
 * Set page details in inches.
*/
Worksheet.prototype.setPageMargin = function (margin) {
	this._margin = _.defaults(margin, {
		left: 0.7,
		right: 0.7,
		top: 0.75,
		bottom: 0.75,
		header: 0.3,
		footer: 0.3
	});
	return this;
};

/**
 * http://www.datypic.com/sc/ooxml/t-ssml_ST_Orientation.html
 *
 * Can be one of 'portrait' or 'landscape'.
 *
 * @param {String} orientation
 */
Worksheet.prototype.setPageOrientation = function (orientation) {
	this._orientation = orientation;
	return this;
};

/**
 * Set rows to repeat for print
 *
 * @param {int|[int, int]} params - number of rows to repeat from the top | [first, last] repeat rows
 */
Worksheet.prototype.setPrintTitleTop = function (params) {
	this._printTitles = this._printTitles || {};

	if (_.isObject(params)) {
		this._printTitles.topFrom = params[0];
		this._printTitles.topTo = params[1];
	} else {
		this._printTitles.topFrom = 0;
		this._printTitles.topTo = params - 1;
	}
	return this;
};

/**
 * Set columns to repeat for print
 *
 * @param {int|[int, int]} params - number of columns to repeat from the left | [first, last] repeat columns
 */
Worksheet.prototype.setPrintTitleLeft = function (params) {
	this._printTitles = this._printTitles || {};

	if (_.isObject(params)) {
		this._printTitles.leftFrom = params[0];
		this._printTitles.leftTo = params[1];
	} else {
		this._printTitles.leftFrom = 0;
		this._printTitles.leftTo = params - 1;
	}
	return this;
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

/**
 * Merge cells in given range
 *
 * @param cell1 - A1, A2...
 * @param cell2 - A2, A3...
 */
Worksheet.prototype.mergeCells = function (cell1, cell2) {
	this.mergedCells.push([cell1, cell2]);
	return this;
};

Worksheet.prototype.setHyperlink = function (hyperlink) {
	hyperlink.objectId = _.uniqueId('hyperlink');
	this.relations.addRelation({
		objectId: hyperlink.objectId,
		target: hyperlink.location,
		targetMode: hyperlink.targetMode || 'External'
	}, 'hyperlink');
	this.hyperlinks.push(hyperlink);
	return this;
};

Worksheet.prototype.setAttribute = function (name, value) {
	this.sheetView.setAttribute(name, value);
	return this;
};

//Add froze pane
Worksheet.prototype.freeze = function (column, row, cell, activePane) {
	this.sheetView.freeze(column, row, cell, activePane);
	return this;
};

//Add split pane
Worksheet.prototype.split = function (x, y, cell, activePane) {
	this.sheetView.split(x, y, cell, activePane);
	return this;
};

Worksheet.prototype._prepare = function () {
	var rowIndex;
	var len;
	var maxX = 0;
	var row;

	this._prepareTables();

	this.preparedData = [];
	this.preparedRows = [];

	for (rowIndex = 0, len = this.data.length; rowIndex < len; rowIndex++) {
		row = prepareRow(this, rowIndex);

		if (row) {
			if (row.length > maxX) {
				maxX = row.length;
			}
		}
	}

	this.data = null;
	this.rows = null;

	this.maxX = maxX;
	this.maxY = this.preparedData.length;
};

Worksheet.prototype._prepareTables = function () {
	var data = this.data;

	_.forEach(this.tables, function (table) {
		table._prepare(data);
	});
};

Worksheet.prototype._export = function (canStream) {
	if (canStream) {
		return new WorksheetStream({
			worksheet: this
		});
	} else {
		return exportBeforeRows(this) +
			exportData(this.preparedData, this.preparedRows, this.timezoneOffset) +
			exportAfterRows(this);
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
		while (this.index < this.len && !stop) {
			while (this.index < this.len && s.length < size) {
				s += exportRow(worksheet.preparedData[this.index], this.index,
					worksheet.preparedRows, worksheet.timezoneOffset);
				worksheet.preparedData[this.index] = null;
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

function prepareRow(worksheet, rowIndex) {
	var styleSheet = worksheet.common.styleSheet;
	var row = worksheet.rows[rowIndex];
	var dataRow = worksheet.data[rowIndex];
	var preparedRow = [];
	var rowStyle = null;
	var column;
	var colIndex;
	var value;
	var cellValue;
	var cellStyle;
	var cellType;
	var colSpan;
	var rowSpan;

	if (dataRow) {
		if (dataRow.data) {
			row = row || {};
			row.height = dataRow.height || row.height;
			row.style = dataRow.style || row.style;
			row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;

			dataRow = dataRow.data;
		}
		if (row) {
			rowStyle = row.style || null;
			row.style = styleSheet.cells.getId(row.style);
		}

		for (colIndex = 0; colIndex < dataRow.length; colIndex++) {
			column = worksheet.columns[colIndex];
			value = dataRow[colIndex];

			if (_.isNil(value)) {
				continue;
			}

			if (value && typeof value === 'object') {
				cellStyle = value.style || null;
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

				if (value.hyperlink) {
					insertHyperlink(worksheet, colIndex, rowIndex, value.hyperlink);
				}
				if (value.image) {
					insertImage(worksheet, colIndex, rowIndex, value.image);
				}
				if (value.colspan || value.rowspan) {
					colSpan = (value.colspan || 1) - 1;
					rowSpan = (value.rowspan || 1) - 1;

					worksheet.mergeCells({c: colIndex + 1, r: rowIndex + 1},
						{c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan});
					insertEmptyCells(worksheet, dataRow, colIndex, rowIndex, colSpan, rowSpan);
				}
			} else {
				cellValue = value;
				cellStyle = null;
				cellType = null;
			}

			if (cellStyle === null) {
				if (rowStyle !== null) {
					cellStyle = rowStyle;
				} else if (column && column.style) {
					cellStyle = column.style;
				}
			}

			if (!cellType) {
				if (column && column.type) {
					cellType = column.type;
				} else if (typeof cellValue === 'number') {
					cellType = 'number';
				} else {
					cellType = 'string';
				}
			}

			if (cellType === 'string') {
				cellValue = worksheet.common.sharedStrings.addString(cellValue);
			}

			preparedRow[colIndex] = {
				value: cellValue,
				style: styleSheet.cells.getId(cellStyle),
				type: cellType
			};
		}
	}

	worksheet.preparedData[rowIndex] = preparedRow;
	worksheet.preparedRows[rowIndex] = row;

	return preparedRow;
}

function insertHyperlink(worksheet, colIndex, rowIndex, hyperlink) {
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
	worksheet.setHyperlink({
		cell: {c: colIndex + 1, r: rowIndex + 1},
		location: location,
		targetMode: targetMode,
		tooltip: tooltip
	});
}

function insertImage(worksheet, colIndex, rowIndex, image) {
	var config;

	if (typeof image === 'string' || image.data) {
		worksheet.setImage(image, {c: colIndex + 1, r: rowIndex + 1});
	} else {
		config = image.config || {};
		config.cell = {c: colIndex + 1, r: rowIndex + 1};

		worksheet.setImage(image.image, config);
	}
}

function insertEmptyCells(worksheet, dataRow, colIndex, rowIndex, colSpan, rowSpan) {
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
			row = worksheet.data[rowIndex + i + 1];

			if (!row) {
				row = [];
				worksheet.data[rowIndex + i + 1] = row;
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
}

function exportBeforeRows(worksheet) {
	return getXMLBegin() +
		exportDimension(worksheet.maxX, worksheet.maxY) +
		worksheet.sheetView._export() +
		exportColumns(worksheet.columns) +
		'<sheetData>';
}

function exportAfterRows(worksheet) {
	return '</sheetData>' +
		// 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
		// with Microsoft Excel (2007, 2013)
		exportMergeCells(worksheet.mergedCells) +
		exportHyperlinks(worksheet.relations, worksheet.hyperlinks) +
		exportPageMargins(worksheet._margin) +
		exportPageSetup(worksheet._orientation) +
		exportHeaderFooter(worksheet.headers, worksheet.footers) +
		exportTables(worksheet.relations, worksheet.tables) +
		// the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
		// to issue with Microsoft Excel (2007, 2013)
		exportDrawings(worksheet.relations, worksheet.drawings) +
		getXMLEnd();
}

function exportData(data, rows, timezoneOffset) {
	var children = '';
	var dataRow;

	for (var i = 0, len = data.length; i < len; i++) {
		dataRow = data[i];
		children += exportRow(dataRow, i, rows, timezoneOffset);
		data[i] = null;
	}
	return children;
}

function exportRow(dataRow, rowIndex, rows, timezoneOffset) {
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
			if (value.style !== null) {
				attrs += ' s="' + value.style + '"';
			}

			switch (value.type) {
				case 'number':
					rowChildren[colIndex] = '<c' + attrs + '><v>' + value.value + '</v></c>';
					break;
				case 'date':
				case 'time':
					rowChildren[colIndex] = '<c' + attrs + '><v>' +
						(25569.0 + (value.value - timezoneOffset) / (60 * 60 * 24 * 1000)) +
						'</v></c>';
					break;
				case 'formula':
					rowChildren[colIndex] = '<c' + attrs + '><f>' + _.escape(value.value) + '</f></c>';
					break;
				case 'string':
					rowChildren[colIndex] = '<c' + attrs + ' t="s"><v>' + value.value + '</v></c>';
					break;
				case 'empty':
					rowChildren[colIndex] = '<c' + attrs + '/>';
					break;
			}
		}
	}

	return '<row' + getRowAttributes(rows[rowIndex], rowIndex) + '>' +
		rowChildren.join('') +
	'</row>';
}

function getRowAttributes(row, rowIndex) {
	var attributes = ' r="' + (rowIndex + 1) + '"';

	if (row) {
		if (row.height !== undefined) {
			attributes += ' customHeight="1"';
			attributes += ' ht="' + row.height + '"';
		}
		if (row.style !== undefined) {
			attributes += ' customFormat="1"';
			attributes += ' s="' + row.style + '"';
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
			if (column.style !== undefined) {
				attributes.push(['style', column.style]);
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

function exportMergeCells(mergedCells) {
	if (mergedCells.length > 0) {
		var children = _.map(mergedCells, function (mergeCell) {
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
				['count', mergedCells.length]
			],
			children: children
		});
	}
	return '';
}

function exportHyperlinks(relations, hyperlinks) {
	if (hyperlinks.length > 0) {
		var children = _.map(hyperlinks, function (hyperlink) {
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

function exportTables(relations, tables) {
	if (tables.length > 0) {
		var children = _.map(tables, function (table) {
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
				['count', tables.length]
			],
			children: children
		});
	}
	return '';
}

function exportDrawings(relations, drawings) {
	if (drawings) {
		return toXMLString({
			name: 'drawing',
			attributes: [
				['r:id', relations.getRelationshipId(drawings)]
			]
		});
	}
	return '';
}

function getXMLBegin() {
	return util.xmlPrefix + '<worksheet xmlns="' + util.schemas.spreadsheetml +
		'" xmlns:r="' + util.schemas.relationships + '" xmlns:mc="' + util.schemas.markupCompat + '">';
}

function getXMLEnd() {
	return '</worksheet>';
}

module.exports = Worksheet;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":2,"./drawings":7,"./relationshipManager":11,"./sheetView":13,"./table":26,"./util":27,"stream":1}]},{},[8])(8)
});