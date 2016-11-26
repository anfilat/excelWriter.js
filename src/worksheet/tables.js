'use strict';

var _ = require('lodash');
var Table = require('../table');
var toXMLString = require('../XMLString');

module.exports = {
	init: function () {
		this._tables = [];
	},
	methods: {
		addTable: function (config) {
			var table = new Table(this, config);

			this.common.addTable(table);
			this.relations.addRelation(table, 'table');
			this._tables.push(table);

			return table;
		},
		_prepareTables: function () {
			var data = this.data;

			_.forEach(this._tables, function (table) {
				table._prepare(data);
			});
		},
		_exportTables: function () {
			var relations = this.relations;

			if (this._tables.length > 0) {
				var children = _.map(this._tables, function (table) {
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
						['count', this._tables.length]
					],
					children: children
				});
			}
			return '';
		}
	}
};
