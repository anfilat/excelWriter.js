'use strict';

const createTable = require('../table');
const toXMLString = require('../XMLString');

function Tables(worksheet, common, relations) {
	this.worksheet = worksheet;
	this.common = common;
	this.relations = relations;
	this.tables = [];
}

Tables.prototype = {
	add(config) {
		const {outerTable, table} = createTable(this.worksheet.outerWorksheet, this.common, config);

		this.common.addTable(table);
		this.relations.add(table, 'table');
		this.tables.push(table);

		return outerTable;
	},
	prepare() {
		this.tables.forEach(table => {
			table.prepare(this.worksheet.data);
		});
	},
	save() {
		if (this.tables.length) {
			const children = this.tables.map(
				table => toXMLString({
					name: 'tablePart',
					attributes: [
						['r:id', this.relations.getId(table)]
					]
				})
			);

			return toXMLString({
				name: 'tableParts',
				attributes: [
					['count', this.tables.length]
				],
				children
			});
		}
		return '';
	}
};

module.exports = Tables;
