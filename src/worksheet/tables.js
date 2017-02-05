'use strict';

const _ = require('lodash');
const Table = require('../table');
const toXMLString = require('../XMLString');

module.exports = SuperClass => class MergedCells extends SuperClass {
	constructor(workbook, config) {
		super(workbook, config);
		this._tables = [];
	}
	addTable(config) {
		const table = new Table(this, config);

		this.common.addTable(table);
		this.relations.addRelation(table, 'table');
		this._tables.push(table);

		return table;
	}
	_prepareTables() {
		const data = this.data;

		_.forEach(this._tables, table => {
			table._prepare(data);
		});
	}
	_saveTables() {
		const relations = this.relations;

		if (this._tables.length > 0) {
			const children = _.map(this._tables,
				table => toXMLString({
					name: 'tablePart',
					attributes: [
						['r:id', relations.getRelationshipId(table)]
					]
				})
			);

			return toXMLString({
				name: 'tableParts',
				attributes: [
					['count', this._tables.length]
				],
				children
			});
		}
		return '';
	}
};
