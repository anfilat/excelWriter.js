'use strict';

const Worksheet = require('./worksheet');

function createWorksheet(outerWorkbook, common, config) {
	const worksheet = new Worksheet(common, config);

	const outerWorksheet = {
		end() {
			return outerWorkbook;
		},
		setActive() {
			worksheet.common.setActiveWorksheet(worksheet);
			return this;
		},
		setVisible() {
			this.setState('visible');
			return this;
		},
		setHidden() {
			this.setState('hidden');
			return this;
		},
		/**
		 * //http://www.datypic.com/sc/ooxml/t-ssml_ST_SheetState.html
		 * @param state - visible | hidden | veryHidden
		 */
		setState(state) {
			worksheet.setState(state);
			return this;
		},
		getState() {
			return worksheet.getState();
		},
		setRows(startRow, rows) {
			worksheet.setRows(startRow, rows);
			return this;
		},
		setRow(rowIndex, row) {
			worksheet.setRow(rowIndex, row);
			return this;
		},
		setColumns(startColumn, columns) {
			worksheet.setColumns(startColumn, columns);
			return this;
		},
		setColumn(columnIndex, column) {
			worksheet.setColumn(columnIndex, column);
			return this;
		},
		setData(offset, data) {
			worksheet.setData(offset, data);
			return this;
		},
		setAttribute(name, value) {
			worksheet.sheetView.setAttribute(name, value);
			return this;
		},
		freeze(col, row, cell, activePane) {
			worksheet.sheetView.freeze(col, row, cell, activePane);
			return this;
		},
		split(x, y, cell, activePane) {
			worksheet.sheetView.split(x, y, cell, activePane);
			return this;
		},
		setHyperlink(hyperlink) {
			worksheet.hyperlinks.setHyperlink(hyperlink);
			return this;
		},
		mergeCells(cell1, cell2) {
			worksheet.mergedCells.mergeCells(cell1, cell2);
			return this;
		},
		setImage(image, config) {
			worksheet.drawings.setImage(image, config);
			return this;
		},
		setImageOneCell(image, config) {
			worksheet.drawings.setImageOneCell(image, config);
			return this;
		},
		setImageAbsolute(image, config) {
			worksheet.drawings.setImageAbsolute(image, config);
			return this;
		},
		addTable(config) {
			return worksheet.tables.addTable(config);
		},
		setHeader(headers) {
			worksheet.setHeader(headers);
			return this;
		},
		setFooter(footers) {
			worksheet.setFooter(footers);
			return this;
		},
		setPageMargin(margin) {
			worksheet.setPageMargin(margin);
			return this;
		},
		setPageOrientation(orientation) {
			worksheet.setPageOrientation(orientation);
			return this;
		},
		setPrintTitleTop(params) {
			worksheet.setPrintTitleTop(params);
			return this;
		},
		setPrintTitleLeft(params) {
			worksheet.setPrintTitleLeft(params);
			return this;
		}
	};
	worksheet.outerWorksheet = outerWorksheet;
	return {
		outerWorksheet,
		worksheet
	};
}

module.exports = createWorksheet;
