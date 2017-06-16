'use strict';

const Table = require('./table');

function createTable(outerWorksheet, common, config) {
	const table = new Table(config, common);

	const outerTable = {
		end() {
			return outerWorksheet;
		},
		setReferenceRange(beginCell, endCell) {
			table.setReferenceRange(beginCell, endCell);
			return this;
		},
		addTotalRow(totalRow) {
			table.addTotalRow(totalRow);
			return this;
		},
		setTheme(theme) {
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
		setSortState(state) {
			table.setSortState(state);
		}
	};
	return {
		outerTable,
		table
	};
}

module.exports = createTable;
