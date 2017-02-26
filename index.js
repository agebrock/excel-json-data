#!/usr/bin/env node

var XLSX = require('xlsx');
var _ = require('lodash');
var util = require('util');
var fs = require('fs');

var yargs = require('yargs')
    .usage('Usage: $0 <file> [options]')
    .option('o', {
           alias: 'output',
           describe: 'write result to given filepath',
           type: 'string'
       })
    .option('c', {
           alias: 'colors',
           describe: 'colorful stdout',
           type: 'boolean'
       })
    .example('$0 foo.xlsx', 'convert given file to json')
    .help('h')
    .alias('h', 'help')
    .epilog('copyright 2017 Christoph.Hagenbrock@googlemail.com');

   var  argv = yargs.argv;

    //var argv = yargs.args;

if(argv._.length !== 1){
  console.log('excel source file required');
 return  yargs.showHelp();
}


var workbook = XLSX.readFile(argv._[0]);


var objUtil = {

    set: function (name, value, context) {
        context = context || myGlobal;

        var parts = name.split('.');
        var p = parts.pop();
        var obj = this._get(parts, true, null, context);
        return obj && p ? (obj[p] = value) : undefined; // Object
    },

    _get: function (parts, create, defaultValue, context) {
        context = context || myGlobal;
        create = create || false;

        defaultValue = (defaultValue === false || defaultValue) ? defaultValue : null;
        try {
            for (var i = 0; i < parts.length; i++) {
                var p = parts[i];
                if (!(p in context)) {
                    if (create) {
                        context[p] = {};
                    } else {
                        return defaultValue;
                    }
                }
                context = context[p];
            }
            return context; // mixed
        } catch (e) {
            return defaultValue;
        }
    }
};




var result = {};
    var collections = {};
    function toJson(workbook) {


        workbook.SheetNames.forEach(function (sheetName) {
            //@todo: this is wrong and works by accident !!
            var roa = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[ sheetName ],{raw:true});

          if (roa.length > 0) {
                collections[ sheetName ] = roa;
            }
        });
        Object.keys(collections).forEach(function (sheetName) {
            var metaTable = getMetaTable(collections, sheetName);
            if (isObjectCollection(sheetName)) {
                result[ sheetName ] = handleMetaTable(collections, metaTable, sheetName);
            }
        });
        Object.keys(result).forEach(function (sheetName) {
            var metaTable = getMetaTable(collections, sheetName);
            if (isObjectCollection(sheetName)) {
                result[ sheetName ] = handleRelations(result, metaTable, sheetName);
            }
        });
        return result;
    }

    var types = [
      {
        type: 'set',
        fn: function(target, value) {
          if (Array.isArray(value)) {
            var objValue = {};
            value.forEach(function(key) {
              objValue[key] = true;
            });
            value = objValue;
          }
          return value;
        }
      },
      {
        type: 'number',
        fn: function(target, value) {
          return parseFloat(value);
        }
      },
      {
        type: 'string',
        fn: function(target, value) {
          return value.toString();
        }
      },
      {
        type: 'hasMany',
        fn: function(target, value, collections, mapConfig) {
          var relationTableName = mapConfig.value || mapConfig.name,
            relationTable = collections[relationTableName];

          if (value) {
            return _.findWhere(relationTable, value);
          }
          return relationTable;
        }
      }
    ];

    var map = {
      getName: function(nameA, nameB) {
        return (nameA) ? nameA : nameB;
      }
    };

    function compileName(name, jsonName, key) {
      var rx;
      if (!jsonName) {
        return name;
      }
      rx = new RegExp(name);
      return key.replace(rx, jsonName);
    }

    function compileValue(value) {
      if (value && value.trim) {
        value = value.trim();
      }

      if (value && value.indexOf && value.indexOf('{') === 0) {
        return JSON.parse(value);
      }
      if (value === 'true') {
        return true;
      }

      if (value === 'false') {
        return false;
      }
      return value;
    }

    function transformItem(collections, meta, obj, value, key) {
      var logic;
      var mapConfig;

      if (key !== null) {
        mapConfig = _.find(meta, function(config) {
          if (config.name === undefined) {
            return false;
          }
          return new RegExp(config.name).test(key);
        });
      }

      if (mapConfig) {
        var query = _.omit({
          type: mapConfig.type,
          action: mapConfig.action
        }, _.isEmpty);

        key = compileName(mapConfig.name, mapConfig.jsonName, key);

        if (!_.isEmpty(query)) {
          logic = _.findWhere(types, query);
        }
        if (logic) {
          value = logic.fn(obj[key], value, collections, mapConfig);
        }

      }
      value = compileValue(value);
      objUtil.set(key, value, obj);
    }

    function handleMetaTable(collections, meta, sheetName) {
      var table = collections[sheetName];
      var boundTransformItem = transformItem.bind(null, collections, meta);
      return _.map(table, function(item) {

        return _.transform(item, boundTransformItem);
      });
    }

    function handleRelations(collections, meta, sheetName) {
      var table = collections[sheetName];
      var relations = _.where(meta, {type: 'relation'});

      relations = relations.map(function(relation) {
        var collectionName = relation.name;
        var key = relation.jsonName || relation.name;
        return function(item) {
          if (!item.id) {
            return null;
          }

          var query = {};
          query['id_' + sheetName] = item.id;
          item[key] = _.where(collections[collectionName], query);
        }
      });
      return _.forEach(table, function(item) {
        relations.forEach(function(relation) {
          relation(item);
        });
      });
    }

    function isObjectCollection(name) {
      return name.indexOf('.') === -1;
    }

    function getMetaTable(tableCollection, name) {
      var metaTableName = name + '.meta';
      if (tableCollection[metaTableName]) {
        return tableCollection[metaTableName];
      } else {
        return [];
      }
    }




    function parse_object(obj, path) {
        var sign = '.';
        if (path == undefined) {
            path = sign;
        }

        var type = _.getType(obj);
        var scalar = (type == "number" || type == "string" || type == "boolean" || type == "null");

        if (type == "array" || type == "object") {
            var d = {};

            if (type == 'array') {
                //  path = path.substr(0, path.length - 1);
            }
            for (var i in obj) {
                var value = i + sign;
                if (type == 'array') {
                    //   value = '[' + i + ']' + sign;
                }
                var newD = parse_object(obj[i], path + value);
                _.extend(d, newD);
            }

            return d;
        }

        else if (scalar) {
            var d = {};
            var endPath = path.substr(1, path.length - 2);
            d[endPath] = obj;
            return d;
        }

        // ?
        else return {};
    }


    function toObj(start, value, key) {
        var ref = start || {};

        var paths = key.split('.').map(function (keypart) {
            var numbericValue = parseInt(keypart);
            if (!isNaN(numbericValue)) {
                keypart = numbericValue;
            }
            return keypart;
        });

        paths.forEach(function (val, index) {
            var nextPath = paths[index + 1],
                isArray = _.isNumber(nextPath);

            if (_.isDefined(nextPath) && !_.isDefined(start[val])) {
                if (isArray) {
                    start[val] = [];
                } else {
                    start[val] = {};
                }
            }

            if (!_.isDefined(nextPath)) {
                start[val] = value;
            }

            start = start[val];
        });
        return ref;
    }

    /**
     * Returns a lowercase string based on the primitive type of a given value.
     *
     * @param value
     * @returns {undefined|null|string|nan|number|boolean|date|function|array|object}
     */
    function getType(value) {
        var type;
        if (_.isUndefined(value)) {
            type = 'undefined';
        } else if (_.isNull(value)) {
            type = 'null';
        } else if (_.isString(value)) {
            type = 'string';
        } else if (_.isNaN(value)) {
            type = 'nan';
        } else if (_.isNumber(value)) {
            type = 'number';
        } else if (_.isBoolean(value)) {
            type = 'boolean';
        } else if (_.isDate(value)) {
            type = 'date';
        } else if (_.isRegExp(value)) {
            type = 'regexp';
        } else if (_.isFunction(value)) {
            type = 'function';
        } else if (_.isArray(value)) {
            type = 'array';
        } else if (_.isObject(value)) {
            type = 'object';
        }
        return type;
    }

    /**
     *
     * Checks a value is undefined or null
     *
     * @param value
     * @returns {boolean}
     */
    function isDefined(value) {
        return getType(value) !== 'null' && getType(value) !== 'undefined';
    }

    _.mixin({getType: getType, isDefined: isDefined});

    var result = toJson(workbook);
    if(argv.colors && !argv.output){
      console.log(util.inspect(result,{colors:true, depth:5}));
    }else if(!argv.output){
      console.log(result);
    }

    if(argv.output){
      fs.writeFileSync(argv.output, JSON.stringify(result));
    }
