'use strict';

const Print = require('./print');
const Table = require('../table');
const toXMLString = require('../XMLString');

class Tables extends Print {
	constructor() {
		super();
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
		this._tables.forEach(table => {
			table._prepare(this.data);
		});
	}
	_saveTables() {
		if (this._tables.length > 0) {
			const children = this._tables.map(
				table => toXMLString({
					name: 'tablePart',
					attributes: [
						['r:id', this.relations.getRelationshipId(table)]
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
}

module.exports = Tables;
