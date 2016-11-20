'use strict';

var _ = require('lodash');
var Table = require('../table');
var toXMLString = require('../XMLString');

function Tables(worksheet) {
	this.worksheet = worksheet;
	this.common = worksheet.common;
	this.relations = worksheet.relations;

	this.tables = [];
}

Tables.prototype.add = function (config) {
	var table = new Table(this.worksheet, config);

	this.common.addTable(table);
	this.relations.addRelation(table, 'table');
	this.tables.push(table);

	return table;
};

Tables.prototype._prepare = function () {
	var data = this.worksheet.data;

	_.forEach(this.tables, function (table) {
		table._prepare(data);
	});
};

Tables.prototype._export = function () {
	var relations = this.relations;

	if (this.tables.length > 0) {
		var children = _.map(this.tables, function (table) {
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
				['count', this.tables.length]
			],
			children: children
		});
	}
	return '';
};

module.exports = Tables;
