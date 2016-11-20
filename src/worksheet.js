'use strict';

var Readable = require('stream').Readable;
var _ = require('lodash');
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
		row = prepareRow(this, rowIndex, this.timezoneOffset);

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
			exportData(this.preparedData, this.preparedRows) +
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
				s += exportRow(worksheet.preparedData[this.index], this.index, worksheet.preparedRows);
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

function prepareRow(worksheet, rowIndex, timezoneOffset) {
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
	var cellFormula;
	var isString;
	var colSpan;
	var rowSpan;
	var date;

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

			cellFormula = null;
			isString = false;
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
				isString = true;
			} else if (cellType === 'date' || cellType === 'time') {
				date = 25569.0 + (cellValue - timezoneOffset) / (60 * 60 * 24 * 1000);
				if (_.isFinite(date)) {
					cellValue = date;
				} else {
					cellValue = worksheet.common.sharedStrings.addString(String(cellValue));
					isString = true;
				}
			} else if (cellType === 'formula') {
				cellValue = null;
				cellFormula = _.escape(value.value);
			}

			preparedRow[colIndex] = {
				value: cellValue,
				formula: cellFormula,
				style: styleSheet.cells.getId(cellStyle),
				isString: isString
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

function exportData(data, rows) {
	var children = '';
	var dataRow;

	for (var i = 0, len = data.length; i < len; i++) {
		dataRow = data[i];
		children += exportRow(dataRow, i, rows);
		data[i] = null;
	}
	return children;
}

function exportRow(dataRow, rowIndex, rows) {
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
			if (value.style) {
				attrs += ' s="' + value.style + '"';
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
