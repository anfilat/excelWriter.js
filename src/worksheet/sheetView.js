'use strict';

/**
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sheetview.aspx
 */
const _ = require('lodash');
const util = require('../util');
const toXMLString = require('../XMLString');

module.exports = SuperClass => class SheetView extends SuperClass {
	constructor(workbook, config) {
		super(workbook, config);
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

		_.forEach(config, (value, name) => {
			if (name === 'freeze') {
				this.freeze(value.col, value.row, value.cell, value.activePane);
			} else if (name === 'split') {
				this.split(value.x, value.y, value.cell, value.activePane);
			} else if (this._attributes[name]) {
				this._attributes[name].value = value;
			}
		});
	}
	setAttribute(name, value) {
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
	freeze(col, row, cell, activePane) {
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
	split(x, y, cell, activePane) {
		this._pane = {
			state: 'split',
			xSplit: x * 20,
			ySplit: y * 20,
			topLeftCell: util.canonCell(cell) || 'A1',
			activePane: activePane || 'bottomRight'
		};
		return this;
	}
	_saveSheetView() {
		const attributes = [
			['workbookViewId', 0]
		];

		_.forEach(this._attributes, (attr, name) => {
			let value = attr.value;

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
					attributes,
					children: [
						savePane(this._pane)
					]
				})
			]
		});
	}
};

function savePane(pane) {
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
