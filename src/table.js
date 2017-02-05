'use strict';

const _ = require('lodash');
const util = require('./util');
const toXMLString = require('./XMLString');

class Table {
	constructor(worksheet, config) {
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
	end() {
		return this.worksheet;
	}
	setReferenceRange(beginCell, endCell) {
		this.beginCell = util.canonCell(beginCell);
		this.endCell = util.canonCell(endCell);
		return this;
	}
	addTotalRow(totalRow) {
		this.totalRow = totalRow;
		this.totalsRowCount = 1;
		return this;
	}
	setTheme(theme) {
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
	setSortState(state) {
		this.sortState = state;
		return this;
	}
	_prepare(worksheetData) {
		const SUB_TOTAL_FUNCTIONS = ['average', 'countNums', 'count', 'max', 'min', 'stdDev', 'sum', 'var'];
		const SUB_TOTAL_NUMS = [101, 102, 103, 104, 105, 107, 109, 110];

		if (this.totalRow) {
			const tableName = this.name;
			const beginCell = util.letterToPosition(this.beginCell);
			const endCell = util.letterToPosition(this.endCell);
			const firstRow = beginCell.y - 1;
			const firstColumn = beginCell.x - 1;
			const lastRow = endCell.y - 1;
			const headerRow = worksheetData[firstRow] || [];
			let totalRow = worksheetData[lastRow + 1];

			if (!totalRow) {
				totalRow = [];
				worksheetData[lastRow + 1] = totalRow;
			}

			_.forEach(this.totalRow, (cell, i) => {
				let headerValue = headerRow[firstColumn + i];
				let funcIndex;

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
							value: `SUBTOTAL(${SUB_TOTAL_NUMS[funcIndex]},${tableName}[${headerValue}])`,
							type: 'formula'
						};
					}
				}
			});
		}
	}
	//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.table.aspx
	_save() {
		const attributes = [
			['id', this.tableId],
			['name', this.name],
			['displayName', this.displayName]
		];
		const children = [];
		const end = util.letterToPosition(this.endCell);
		const ref = this.beginCell + ':' + util.positionToLetter(end.x, end.y + this.totalsRowCount);

		attributes.push(['ref', ref]);
		attributes.push(['totalsRowCount', this.totalsRowCount]);
		attributes.push(['headerRowCount', this.headerRowCount]);

		children.push(saveAutoFilter(this.beginCell, this.endCell));
		children.push(saveTableColumns(this.totalRow));
		children.push(saveTableStyleInfo(this.common, this.themeStyle));

		return toXMLString({
			name: 'table',
			ns: 'spreadsheetml',
			attributes,
			children
		});
	}
}

function saveAutoFilter(beginCell, endCell) {
	return toXMLString({
		name: 'autoFilter',
		attributes: ['ref', beginCell + ':' + endCell]
	});
}

function saveTableColumns(totalRow) {
	const attributes = [
		['count', totalRow.length]
	];
	const children = _.map(totalRow, (cell, index) => {
		const attributes = [
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
			attributes
		});
	});

	return toXMLString({
		name: 'tableColumns',
		attributes,
		children
	});
}

function saveTableStyleInfo(common, themeStyle) {
	const attributes = [
		['name', themeStyle]
	];
	const format = common.styles.tables.get(themeStyle);
	let isRowStripes = false;
	let isColumnStripes = false;
	let isFirstColumn = false;
	let isLastColumn = false;

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
		attributes
	});
}

module.exports = Table;
