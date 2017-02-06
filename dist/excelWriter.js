/*
excelWriter - Javascript module for generating Excel files
<https://github.com/anfilat/excelWriter.js>
(c) 2017 Andrey Filatkin
Dual licenced under the MIT license or GPLv3
*/
(function(f){if(typeof exports==="object"&&typeof module!=="undefined"){module.exports=f()}else if(typeof define==="function"&&define.amd){define([],f)}else{var g;if(typeof window!=="undefined"){g=window}else if(typeof global!=="undefined"){g=global}else if(typeof self!=="undefined"){g=self}else{g=this}g.excelWriter = f()}})(function(){var define,module,exports;return (function e(t,n,r){function s(o,u){if(!n[o]){if(!t[o]){var a=typeof require=="function"&&require;if(!u&&a)return a(o,!0);if(i)return i(o,!0);var f=new Error("Cannot find module '"+o+"'");throw f.code="MODULE_NOT_FOUND",f}var l=n[o]={exports:{}};t[o][0].call(l.exports,function(e){var n=t[o][1][e];return s(n?n:e)},l,l.exports,e,t,n,r)}return n[o].exports}var i=typeof require=="function"&&require;for(var o=0;o<r.length;o++)s(r[o]);return s})({1:[function(require,module,exports){
module.exports = { "default": require("core-js/library/fn/json/stringify"), __esModule: true };
},{"core-js/library/fn/json/stringify":14}],2:[function(require,module,exports){
module.exports = { "default": require("core-js/library/fn/map"), __esModule: true };
},{"core-js/library/fn/map":15}],3:[function(require,module,exports){
module.exports = { "default": require("core-js/library/fn/object/create"), __esModule: true };
},{"core-js/library/fn/object/create":16}],4:[function(require,module,exports){
module.exports = { "default": require("core-js/library/fn/object/define-property"), __esModule: true };
},{"core-js/library/fn/object/define-property":17}],5:[function(require,module,exports){
module.exports = { "default": require("core-js/library/fn/object/get-prototype-of"), __esModule: true };
},{"core-js/library/fn/object/get-prototype-of":18}],6:[function(require,module,exports){
module.exports = { "default": require("core-js/library/fn/object/set-prototype-of"), __esModule: true };
},{"core-js/library/fn/object/set-prototype-of":19}],7:[function(require,module,exports){
module.exports = { "default": require("core-js/library/fn/symbol"), __esModule: true };
},{"core-js/library/fn/symbol":20}],8:[function(require,module,exports){
module.exports = { "default": require("core-js/library/fn/symbol/iterator"), __esModule: true };
},{"core-js/library/fn/symbol/iterator":21}],9:[function(require,module,exports){
"use strict";

exports.__esModule = true;

exports.default = function (instance, Constructor) {
  if (!(instance instanceof Constructor)) {
    throw new TypeError("Cannot call a class as a function");
  }
};
},{}],10:[function(require,module,exports){
"use strict";

exports.__esModule = true;

var _defineProperty = require("../core-js/object/define-property");

var _defineProperty2 = _interopRequireDefault(_defineProperty);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

exports.default = function () {
  function defineProperties(target, props) {
    for (var i = 0; i < props.length; i++) {
      var descriptor = props[i];
      descriptor.enumerable = descriptor.enumerable || false;
      descriptor.configurable = true;
      if ("value" in descriptor) descriptor.writable = true;
      (0, _defineProperty2.default)(target, descriptor.key, descriptor);
    }
  }

  return function (Constructor, protoProps, staticProps) {
    if (protoProps) defineProperties(Constructor.prototype, protoProps);
    if (staticProps) defineProperties(Constructor, staticProps);
    return Constructor;
  };
}();
},{"../core-js/object/define-property":4}],11:[function(require,module,exports){
"use strict";

exports.__esModule = true;

var _setPrototypeOf = require("../core-js/object/set-prototype-of");

var _setPrototypeOf2 = _interopRequireDefault(_setPrototypeOf);

var _create = require("../core-js/object/create");

var _create2 = _interopRequireDefault(_create);

var _typeof2 = require("../helpers/typeof");

var _typeof3 = _interopRequireDefault(_typeof2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

exports.default = function (subClass, superClass) {
  if (typeof superClass !== "function" && superClass !== null) {
    throw new TypeError("Super expression must either be null or a function, not " + (typeof superClass === "undefined" ? "undefined" : (0, _typeof3.default)(superClass)));
  }

  subClass.prototype = (0, _create2.default)(superClass && superClass.prototype, {
    constructor: {
      value: subClass,
      enumerable: false,
      writable: true,
      configurable: true
    }
  });
  if (superClass) _setPrototypeOf2.default ? (0, _setPrototypeOf2.default)(subClass, superClass) : subClass.__proto__ = superClass;
};
},{"../core-js/object/create":3,"../core-js/object/set-prototype-of":6,"../helpers/typeof":13}],12:[function(require,module,exports){
"use strict";

exports.__esModule = true;

var _typeof2 = require("../helpers/typeof");

var _typeof3 = _interopRequireDefault(_typeof2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

exports.default = function (self, call) {
  if (!self) {
    throw new ReferenceError("this hasn't been initialised - super() hasn't been called");
  }

  return call && ((typeof call === "undefined" ? "undefined" : (0, _typeof3.default)(call)) === "object" || typeof call === "function") ? call : self;
};
},{"../helpers/typeof":13}],13:[function(require,module,exports){
"use strict";

exports.__esModule = true;

var _iterator = require("../core-js/symbol/iterator");

var _iterator2 = _interopRequireDefault(_iterator);

var _symbol = require("../core-js/symbol");

var _symbol2 = _interopRequireDefault(_symbol);

var _typeof = typeof _symbol2.default === "function" && typeof _iterator2.default === "symbol" ? function (obj) { return typeof obj; } : function (obj) { return obj && typeof _symbol2.default === "function" && obj.constructor === _symbol2.default && obj !== _symbol2.default.prototype ? "symbol" : typeof obj; };

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

exports.default = typeof _symbol2.default === "function" && _typeof(_iterator2.default) === "symbol" ? function (obj) {
  return typeof obj === "undefined" ? "undefined" : _typeof(obj);
} : function (obj) {
  return obj && typeof _symbol2.default === "function" && obj.constructor === _symbol2.default && obj !== _symbol2.default.prototype ? "symbol" : typeof obj === "undefined" ? "undefined" : _typeof(obj);
};
},{"../core-js/symbol":7,"../core-js/symbol/iterator":8}],14:[function(require,module,exports){
var core  = require('../../modules/_core')
  , $JSON = core.JSON || (core.JSON = {stringify: JSON.stringify});
module.exports = function stringify(it){ // eslint-disable-line no-unused-vars
  return $JSON.stringify.apply($JSON, arguments);
};
},{"../../modules/_core":36}],15:[function(require,module,exports){
require('../modules/es6.object.to-string');
require('../modules/es6.string.iterator');
require('../modules/web.dom.iterable');
require('../modules/es6.map');
require('../modules/es7.map.to-json');
module.exports = require('../modules/_core').Map;
},{"../modules/_core":36,"../modules/es6.map":96,"../modules/es6.object.to-string":101,"../modules/es6.string.iterator":102,"../modules/es7.map.to-json":104,"../modules/web.dom.iterable":107}],16:[function(require,module,exports){
require('../../modules/es6.object.create');
var $Object = require('../../modules/_core').Object;
module.exports = function create(P, D){
  return $Object.create(P, D);
};
},{"../../modules/_core":36,"../../modules/es6.object.create":97}],17:[function(require,module,exports){
require('../../modules/es6.object.define-property');
var $Object = require('../../modules/_core').Object;
module.exports = function defineProperty(it, key, desc){
  return $Object.defineProperty(it, key, desc);
};
},{"../../modules/_core":36,"../../modules/es6.object.define-property":98}],18:[function(require,module,exports){
require('../../modules/es6.object.get-prototype-of');
module.exports = require('../../modules/_core').Object.getPrototypeOf;
},{"../../modules/_core":36,"../../modules/es6.object.get-prototype-of":99}],19:[function(require,module,exports){
require('../../modules/es6.object.set-prototype-of');
module.exports = require('../../modules/_core').Object.setPrototypeOf;
},{"../../modules/_core":36,"../../modules/es6.object.set-prototype-of":100}],20:[function(require,module,exports){
require('../../modules/es6.symbol');
require('../../modules/es6.object.to-string');
require('../../modules/es7.symbol.async-iterator');
require('../../modules/es7.symbol.observable');
module.exports = require('../../modules/_core').Symbol;
},{"../../modules/_core":36,"../../modules/es6.object.to-string":101,"../../modules/es6.symbol":103,"../../modules/es7.symbol.async-iterator":105,"../../modules/es7.symbol.observable":106}],21:[function(require,module,exports){
require('../../modules/es6.string.iterator');
require('../../modules/web.dom.iterable');
module.exports = require('../../modules/_wks-ext').f('iterator');
},{"../../modules/_wks-ext":92,"../../modules/es6.string.iterator":102,"../../modules/web.dom.iterable":107}],22:[function(require,module,exports){
module.exports = function(it){
  if(typeof it != 'function')throw TypeError(it + ' is not a function!');
  return it;
};
},{}],23:[function(require,module,exports){
module.exports = function(){ /* empty */ };
},{}],24:[function(require,module,exports){
module.exports = function(it, Constructor, name, forbiddenField){
  if(!(it instanceof Constructor) || (forbiddenField !== undefined && forbiddenField in it)){
    throw TypeError(name + ': incorrect invocation!');
  } return it;
};
},{}],25:[function(require,module,exports){
var isObject = require('./_is-object');
module.exports = function(it){
  if(!isObject(it))throw TypeError(it + ' is not an object!');
  return it;
};
},{"./_is-object":54}],26:[function(require,module,exports){
var forOf = require('./_for-of');

module.exports = function(iter, ITERATOR){
  var result = [];
  forOf(iter, false, result.push, result, ITERATOR);
  return result;
};

},{"./_for-of":45}],27:[function(require,module,exports){
// false -> Array#indexOf
// true  -> Array#includes
var toIObject = require('./_to-iobject')
  , toLength  = require('./_to-length')
  , toIndex   = require('./_to-index');
module.exports = function(IS_INCLUDES){
  return function($this, el, fromIndex){
    var O      = toIObject($this)
      , length = toLength(O.length)
      , index  = toIndex(fromIndex, length)
      , value;
    // Array#includes uses SameValueZero equality algorithm
    if(IS_INCLUDES && el != el)while(length > index){
      value = O[index++];
      if(value != value)return true;
    // Array#toIndex ignores holes, Array#includes - not
    } else for(;length > index; index++)if(IS_INCLUDES || index in O){
      if(O[index] === el)return IS_INCLUDES || index || 0;
    } return !IS_INCLUDES && -1;
  };
};
},{"./_to-index":84,"./_to-iobject":86,"./_to-length":87}],28:[function(require,module,exports){
// 0 -> Array#forEach
// 1 -> Array#map
// 2 -> Array#filter
// 3 -> Array#some
// 4 -> Array#every
// 5 -> Array#find
// 6 -> Array#findIndex
var ctx      = require('./_ctx')
  , IObject  = require('./_iobject')
  , toObject = require('./_to-object')
  , toLength = require('./_to-length')
  , asc      = require('./_array-species-create');
module.exports = function(TYPE, $create){
  var IS_MAP        = TYPE == 1
    , IS_FILTER     = TYPE == 2
    , IS_SOME       = TYPE == 3
    , IS_EVERY      = TYPE == 4
    , IS_FIND_INDEX = TYPE == 6
    , NO_HOLES      = TYPE == 5 || IS_FIND_INDEX
    , create        = $create || asc;
  return function($this, callbackfn, that){
    var O      = toObject($this)
      , self   = IObject(O)
      , f      = ctx(callbackfn, that, 3)
      , length = toLength(self.length)
      , index  = 0
      , result = IS_MAP ? create($this, length) : IS_FILTER ? create($this, 0) : undefined
      , val, res;
    for(;length > index; index++)if(NO_HOLES || index in self){
      val = self[index];
      res = f(val, index, O);
      if(TYPE){
        if(IS_MAP)result[index] = res;            // map
        else if(res)switch(TYPE){
          case 3: return true;                    // some
          case 5: return val;                     // find
          case 6: return index;                   // findIndex
          case 2: result.push(val);               // filter
        } else if(IS_EVERY)return false;          // every
      }
    }
    return IS_FIND_INDEX ? -1 : IS_SOME || IS_EVERY ? IS_EVERY : result;
  };
};
},{"./_array-species-create":30,"./_ctx":37,"./_iobject":51,"./_to-length":87,"./_to-object":88}],29:[function(require,module,exports){
var isObject = require('./_is-object')
  , isArray  = require('./_is-array')
  , SPECIES  = require('./_wks')('species');

module.exports = function(original){
  var C;
  if(isArray(original)){
    C = original.constructor;
    // cross-realm fallback
    if(typeof C == 'function' && (C === Array || isArray(C.prototype)))C = undefined;
    if(isObject(C)){
      C = C[SPECIES];
      if(C === null)C = undefined;
    }
  } return C === undefined ? Array : C;
};
},{"./_is-array":53,"./_is-object":54,"./_wks":93}],30:[function(require,module,exports){
// 9.4.2.3 ArraySpeciesCreate(originalArray, length)
var speciesConstructor = require('./_array-species-constructor');

module.exports = function(original, length){
  return new (speciesConstructor(original))(length);
};
},{"./_array-species-constructor":29}],31:[function(require,module,exports){
// getting tag from 19.1.3.6 Object.prototype.toString()
var cof = require('./_cof')
  , TAG = require('./_wks')('toStringTag')
  // ES3 wrong here
  , ARG = cof(function(){ return arguments; }()) == 'Arguments';

// fallback for IE11 Script Access Denied error
var tryGet = function(it, key){
  try {
    return it[key];
  } catch(e){ /* empty */ }
};

module.exports = function(it){
  var O, T, B;
  return it === undefined ? 'Undefined' : it === null ? 'Null'
    // @@toStringTag case
    : typeof (T = tryGet(O = Object(it), TAG)) == 'string' ? T
    // builtinTag case
    : ARG ? cof(O)
    // ES3 arguments fallback
    : (B = cof(O)) == 'Object' && typeof O.callee == 'function' ? 'Arguments' : B;
};
},{"./_cof":32,"./_wks":93}],32:[function(require,module,exports){
var toString = {}.toString;

module.exports = function(it){
  return toString.call(it).slice(8, -1);
};
},{}],33:[function(require,module,exports){
'use strict';
var dP          = require('./_object-dp').f
  , create      = require('./_object-create')
  , redefineAll = require('./_redefine-all')
  , ctx         = require('./_ctx')
  , anInstance  = require('./_an-instance')
  , defined     = require('./_defined')
  , forOf       = require('./_for-of')
  , $iterDefine = require('./_iter-define')
  , step        = require('./_iter-step')
  , setSpecies  = require('./_set-species')
  , DESCRIPTORS = require('./_descriptors')
  , fastKey     = require('./_meta').fastKey
  , SIZE        = DESCRIPTORS ? '_s' : 'size';

var getEntry = function(that, key){
  // fast case
  var index = fastKey(key), entry;
  if(index !== 'F')return that._i[index];
  // frozen object case
  for(entry = that._f; entry; entry = entry.n){
    if(entry.k == key)return entry;
  }
};

module.exports = {
  getConstructor: function(wrapper, NAME, IS_MAP, ADDER){
    var C = wrapper(function(that, iterable){
      anInstance(that, C, NAME, '_i');
      that._i = create(null); // index
      that._f = undefined;    // first entry
      that._l = undefined;    // last entry
      that[SIZE] = 0;         // size
      if(iterable != undefined)forOf(iterable, IS_MAP, that[ADDER], that);
    });
    redefineAll(C.prototype, {
      // 23.1.3.1 Map.prototype.clear()
      // 23.2.3.2 Set.prototype.clear()
      clear: function clear(){
        for(var that = this, data = that._i, entry = that._f; entry; entry = entry.n){
          entry.r = true;
          if(entry.p)entry.p = entry.p.n = undefined;
          delete data[entry.i];
        }
        that._f = that._l = undefined;
        that[SIZE] = 0;
      },
      // 23.1.3.3 Map.prototype.delete(key)
      // 23.2.3.4 Set.prototype.delete(value)
      'delete': function(key){
        var that  = this
          , entry = getEntry(that, key);
        if(entry){
          var next = entry.n
            , prev = entry.p;
          delete that._i[entry.i];
          entry.r = true;
          if(prev)prev.n = next;
          if(next)next.p = prev;
          if(that._f == entry)that._f = next;
          if(that._l == entry)that._l = prev;
          that[SIZE]--;
        } return !!entry;
      },
      // 23.2.3.6 Set.prototype.forEach(callbackfn, thisArg = undefined)
      // 23.1.3.5 Map.prototype.forEach(callbackfn, thisArg = undefined)
      forEach: function forEach(callbackfn /*, that = undefined */){
        anInstance(this, C, 'forEach');
        var f = ctx(callbackfn, arguments.length > 1 ? arguments[1] : undefined, 3)
          , entry;
        while(entry = entry ? entry.n : this._f){
          f(entry.v, entry.k, this);
          // revert to the last existing entry
          while(entry && entry.r)entry = entry.p;
        }
      },
      // 23.1.3.7 Map.prototype.has(key)
      // 23.2.3.7 Set.prototype.has(value)
      has: function has(key){
        return !!getEntry(this, key);
      }
    });
    if(DESCRIPTORS)dP(C.prototype, 'size', {
      get: function(){
        return defined(this[SIZE]);
      }
    });
    return C;
  },
  def: function(that, key, value){
    var entry = getEntry(that, key)
      , prev, index;
    // change existing entry
    if(entry){
      entry.v = value;
    // create new entry
    } else {
      that._l = entry = {
        i: index = fastKey(key, true), // <- index
        k: key,                        // <- key
        v: value,                      // <- value
        p: prev = that._l,             // <- previous entry
        n: undefined,                  // <- next entry
        r: false                       // <- removed
      };
      if(!that._f)that._f = entry;
      if(prev)prev.n = entry;
      that[SIZE]++;
      // add to index
      if(index !== 'F')that._i[index] = entry;
    } return that;
  },
  getEntry: getEntry,
  setStrong: function(C, NAME, IS_MAP){
    // add .keys, .values, .entries, [@@iterator]
    // 23.1.3.4, 23.1.3.8, 23.1.3.11, 23.1.3.12, 23.2.3.5, 23.2.3.8, 23.2.3.10, 23.2.3.11
    $iterDefine(C, NAME, function(iterated, kind){
      this._t = iterated;  // target
      this._k = kind;      // kind
      this._l = undefined; // previous
    }, function(){
      var that  = this
        , kind  = that._k
        , entry = that._l;
      // revert to the last existing entry
      while(entry && entry.r)entry = entry.p;
      // get next entry
      if(!that._t || !(that._l = entry = entry ? entry.n : that._t._f)){
        // or finish the iteration
        that._t = undefined;
        return step(1);
      }
      // return step by kind
      if(kind == 'keys'  )return step(0, entry.k);
      if(kind == 'values')return step(0, entry.v);
      return step(0, [entry.k, entry.v]);
    }, IS_MAP ? 'entries' : 'values' , !IS_MAP, true);

    // add [@@species], 23.1.2.2, 23.2.2.2
    setSpecies(NAME);
  }
};
},{"./_an-instance":24,"./_ctx":37,"./_defined":38,"./_descriptors":39,"./_for-of":45,"./_iter-define":57,"./_iter-step":58,"./_meta":62,"./_object-create":63,"./_object-dp":64,"./_redefine-all":76,"./_set-species":79}],34:[function(require,module,exports){
// https://github.com/DavidBruant/Map-Set.prototype.toJSON
var classof = require('./_classof')
  , from    = require('./_array-from-iterable');
module.exports = function(NAME){
  return function toJSON(){
    if(classof(this) != NAME)throw TypeError(NAME + "#toJSON isn't generic");
    return from(this);
  };
};
},{"./_array-from-iterable":26,"./_classof":31}],35:[function(require,module,exports){
'use strict';
var global         = require('./_global')
  , $export        = require('./_export')
  , meta           = require('./_meta')
  , fails          = require('./_fails')
  , hide           = require('./_hide')
  , redefineAll    = require('./_redefine-all')
  , forOf          = require('./_for-of')
  , anInstance     = require('./_an-instance')
  , isObject       = require('./_is-object')
  , setToStringTag = require('./_set-to-string-tag')
  , dP             = require('./_object-dp').f
  , each           = require('./_array-methods')(0)
  , DESCRIPTORS    = require('./_descriptors');

module.exports = function(NAME, wrapper, methods, common, IS_MAP, IS_WEAK){
  var Base  = global[NAME]
    , C     = Base
    , ADDER = IS_MAP ? 'set' : 'add'
    , proto = C && C.prototype
    , O     = {};
  if(!DESCRIPTORS || typeof C != 'function' || !(IS_WEAK || proto.forEach && !fails(function(){
    new C().entries().next();
  }))){
    // create collection constructor
    C = common.getConstructor(wrapper, NAME, IS_MAP, ADDER);
    redefineAll(C.prototype, methods);
    meta.NEED = true;
  } else {
    C = wrapper(function(target, iterable){
      anInstance(target, C, NAME, '_c');
      target._c = new Base;
      if(iterable != undefined)forOf(iterable, IS_MAP, target[ADDER], target);
    });
    each('add,clear,delete,forEach,get,has,set,keys,values,entries,toJSON'.split(','),function(KEY){
      var IS_ADDER = KEY == 'add' || KEY == 'set';
      if(KEY in proto && !(IS_WEAK && KEY == 'clear'))hide(C.prototype, KEY, function(a, b){
        anInstance(this, C, KEY);
        if(!IS_ADDER && IS_WEAK && !isObject(a))return KEY == 'get' ? undefined : false;
        var result = this._c[KEY](a === 0 ? 0 : a, b);
        return IS_ADDER ? this : result;
      });
    });
    if('size' in proto)dP(C.prototype, 'size', {
      get: function(){
        return this._c.size;
      }
    });
  }

  setToStringTag(C, NAME);

  O[NAME] = C;
  $export($export.G + $export.W + $export.F, O);

  if(!IS_WEAK)common.setStrong(C, NAME, IS_MAP);

  return C;
};
},{"./_an-instance":24,"./_array-methods":28,"./_descriptors":39,"./_export":43,"./_fails":44,"./_for-of":45,"./_global":46,"./_hide":48,"./_is-object":54,"./_meta":62,"./_object-dp":64,"./_redefine-all":76,"./_set-to-string-tag":80}],36:[function(require,module,exports){
var core = module.exports = {version: '2.4.0'};
if(typeof __e == 'number')__e = core; // eslint-disable-line no-undef
},{}],37:[function(require,module,exports){
// optional / simple context binding
var aFunction = require('./_a-function');
module.exports = function(fn, that, length){
  aFunction(fn);
  if(that === undefined)return fn;
  switch(length){
    case 1: return function(a){
      return fn.call(that, a);
    };
    case 2: return function(a, b){
      return fn.call(that, a, b);
    };
    case 3: return function(a, b, c){
      return fn.call(that, a, b, c);
    };
  }
  return function(/* ...args */){
    return fn.apply(that, arguments);
  };
};
},{"./_a-function":22}],38:[function(require,module,exports){
// 7.2.1 RequireObjectCoercible(argument)
module.exports = function(it){
  if(it == undefined)throw TypeError("Can't call method on  " + it);
  return it;
};
},{}],39:[function(require,module,exports){
// Thank's IE8 for his funny defineProperty
module.exports = !require('./_fails')(function(){
  return Object.defineProperty({}, 'a', {get: function(){ return 7; }}).a != 7;
});
},{"./_fails":44}],40:[function(require,module,exports){
var isObject = require('./_is-object')
  , document = require('./_global').document
  // in old IE typeof document.createElement is 'object'
  , is = isObject(document) && isObject(document.createElement);
module.exports = function(it){
  return is ? document.createElement(it) : {};
};
},{"./_global":46,"./_is-object":54}],41:[function(require,module,exports){
// IE 8- don't enum bug keys
module.exports = (
  'constructor,hasOwnProperty,isPrototypeOf,propertyIsEnumerable,toLocaleString,toString,valueOf'
).split(',');
},{}],42:[function(require,module,exports){
// all enumerable object keys, includes symbols
var getKeys = require('./_object-keys')
  , gOPS    = require('./_object-gops')
  , pIE     = require('./_object-pie');
module.exports = function(it){
  var result     = getKeys(it)
    , getSymbols = gOPS.f;
  if(getSymbols){
    var symbols = getSymbols(it)
      , isEnum  = pIE.f
      , i       = 0
      , key;
    while(symbols.length > i)if(isEnum.call(it, key = symbols[i++]))result.push(key);
  } return result;
};
},{"./_object-gops":69,"./_object-keys":72,"./_object-pie":73}],43:[function(require,module,exports){
var global    = require('./_global')
  , core      = require('./_core')
  , ctx       = require('./_ctx')
  , hide      = require('./_hide')
  , PROTOTYPE = 'prototype';

var $export = function(type, name, source){
  var IS_FORCED = type & $export.F
    , IS_GLOBAL = type & $export.G
    , IS_STATIC = type & $export.S
    , IS_PROTO  = type & $export.P
    , IS_BIND   = type & $export.B
    , IS_WRAP   = type & $export.W
    , exports   = IS_GLOBAL ? core : core[name] || (core[name] = {})
    , expProto  = exports[PROTOTYPE]
    , target    = IS_GLOBAL ? global : IS_STATIC ? global[name] : (global[name] || {})[PROTOTYPE]
    , key, own, out;
  if(IS_GLOBAL)source = name;
  for(key in source){
    // contains in native
    own = !IS_FORCED && target && target[key] !== undefined;
    if(own && key in exports)continue;
    // export native or passed
    out = own ? target[key] : source[key];
    // prevent global pollution for namespaces
    exports[key] = IS_GLOBAL && typeof target[key] != 'function' ? source[key]
    // bind timers to global for call from export context
    : IS_BIND && own ? ctx(out, global)
    // wrap global constructors for prevent change them in library
    : IS_WRAP && target[key] == out ? (function(C){
      var F = function(a, b, c){
        if(this instanceof C){
          switch(arguments.length){
            case 0: return new C;
            case 1: return new C(a);
            case 2: return new C(a, b);
          } return new C(a, b, c);
        } return C.apply(this, arguments);
      };
      F[PROTOTYPE] = C[PROTOTYPE];
      return F;
    // make static versions for prototype methods
    })(out) : IS_PROTO && typeof out == 'function' ? ctx(Function.call, out) : out;
    // export proto methods to core.%CONSTRUCTOR%.methods.%NAME%
    if(IS_PROTO){
      (exports.virtual || (exports.virtual = {}))[key] = out;
      // export proto methods to core.%CONSTRUCTOR%.prototype.%NAME%
      if(type & $export.R && expProto && !expProto[key])hide(expProto, key, out);
    }
  }
};
// type bitmap
$export.F = 1;   // forced
$export.G = 2;   // global
$export.S = 4;   // static
$export.P = 8;   // proto
$export.B = 16;  // bind
$export.W = 32;  // wrap
$export.U = 64;  // safe
$export.R = 128; // real proto method for `library` 
module.exports = $export;
},{"./_core":36,"./_ctx":37,"./_global":46,"./_hide":48}],44:[function(require,module,exports){
module.exports = function(exec){
  try {
    return !!exec();
  } catch(e){
    return true;
  }
};
},{}],45:[function(require,module,exports){
var ctx         = require('./_ctx')
  , call        = require('./_iter-call')
  , isArrayIter = require('./_is-array-iter')
  , anObject    = require('./_an-object')
  , toLength    = require('./_to-length')
  , getIterFn   = require('./core.get-iterator-method')
  , BREAK       = {}
  , RETURN      = {};
var exports = module.exports = function(iterable, entries, fn, that, ITERATOR){
  var iterFn = ITERATOR ? function(){ return iterable; } : getIterFn(iterable)
    , f      = ctx(fn, that, entries ? 2 : 1)
    , index  = 0
    , length, step, iterator, result;
  if(typeof iterFn != 'function')throw TypeError(iterable + ' is not iterable!');
  // fast case for arrays with default iterator
  if(isArrayIter(iterFn))for(length = toLength(iterable.length); length > index; index++){
    result = entries ? f(anObject(step = iterable[index])[0], step[1]) : f(iterable[index]);
    if(result === BREAK || result === RETURN)return result;
  } else for(iterator = iterFn.call(iterable); !(step = iterator.next()).done; ){
    result = call(iterator, f, step.value, entries);
    if(result === BREAK || result === RETURN)return result;
  }
};
exports.BREAK  = BREAK;
exports.RETURN = RETURN;
},{"./_an-object":25,"./_ctx":37,"./_is-array-iter":52,"./_iter-call":55,"./_to-length":87,"./core.get-iterator-method":94}],46:[function(require,module,exports){
// https://github.com/zloirock/core-js/issues/86#issuecomment-115759028
var global = module.exports = typeof window != 'undefined' && window.Math == Math
  ? window : typeof self != 'undefined' && self.Math == Math ? self : Function('return this')();
if(typeof __g == 'number')__g = global; // eslint-disable-line no-undef
},{}],47:[function(require,module,exports){
var hasOwnProperty = {}.hasOwnProperty;
module.exports = function(it, key){
  return hasOwnProperty.call(it, key);
};
},{}],48:[function(require,module,exports){
var dP         = require('./_object-dp')
  , createDesc = require('./_property-desc');
module.exports = require('./_descriptors') ? function(object, key, value){
  return dP.f(object, key, createDesc(1, value));
} : function(object, key, value){
  object[key] = value;
  return object;
};
},{"./_descriptors":39,"./_object-dp":64,"./_property-desc":75}],49:[function(require,module,exports){
module.exports = require('./_global').document && document.documentElement;
},{"./_global":46}],50:[function(require,module,exports){
module.exports = !require('./_descriptors') && !require('./_fails')(function(){
  return Object.defineProperty(require('./_dom-create')('div'), 'a', {get: function(){ return 7; }}).a != 7;
});
},{"./_descriptors":39,"./_dom-create":40,"./_fails":44}],51:[function(require,module,exports){
// fallback for non-array-like ES3 and non-enumerable old V8 strings
var cof = require('./_cof');
module.exports = Object('z').propertyIsEnumerable(0) ? Object : function(it){
  return cof(it) == 'String' ? it.split('') : Object(it);
};
},{"./_cof":32}],52:[function(require,module,exports){
// check on default Array iterator
var Iterators  = require('./_iterators')
  , ITERATOR   = require('./_wks')('iterator')
  , ArrayProto = Array.prototype;

module.exports = function(it){
  return it !== undefined && (Iterators.Array === it || ArrayProto[ITERATOR] === it);
};
},{"./_iterators":59,"./_wks":93}],53:[function(require,module,exports){
// 7.2.2 IsArray(argument)
var cof = require('./_cof');
module.exports = Array.isArray || function isArray(arg){
  return cof(arg) == 'Array';
};
},{"./_cof":32}],54:[function(require,module,exports){
module.exports = function(it){
  return typeof it === 'object' ? it !== null : typeof it === 'function';
};
},{}],55:[function(require,module,exports){
// call something on iterator step with safe closing on error
var anObject = require('./_an-object');
module.exports = function(iterator, fn, value, entries){
  try {
    return entries ? fn(anObject(value)[0], value[1]) : fn(value);
  // 7.4.6 IteratorClose(iterator, completion)
  } catch(e){
    var ret = iterator['return'];
    if(ret !== undefined)anObject(ret.call(iterator));
    throw e;
  }
};
},{"./_an-object":25}],56:[function(require,module,exports){
'use strict';
var create         = require('./_object-create')
  , descriptor     = require('./_property-desc')
  , setToStringTag = require('./_set-to-string-tag')
  , IteratorPrototype = {};

// 25.1.2.1.1 %IteratorPrototype%[@@iterator]()
require('./_hide')(IteratorPrototype, require('./_wks')('iterator'), function(){ return this; });

module.exports = function(Constructor, NAME, next){
  Constructor.prototype = create(IteratorPrototype, {next: descriptor(1, next)});
  setToStringTag(Constructor, NAME + ' Iterator');
};
},{"./_hide":48,"./_object-create":63,"./_property-desc":75,"./_set-to-string-tag":80,"./_wks":93}],57:[function(require,module,exports){
'use strict';
var LIBRARY        = require('./_library')
  , $export        = require('./_export')
  , redefine       = require('./_redefine')
  , hide           = require('./_hide')
  , has            = require('./_has')
  , Iterators      = require('./_iterators')
  , $iterCreate    = require('./_iter-create')
  , setToStringTag = require('./_set-to-string-tag')
  , getPrototypeOf = require('./_object-gpo')
  , ITERATOR       = require('./_wks')('iterator')
  , BUGGY          = !([].keys && 'next' in [].keys()) // Safari has buggy iterators w/o `next`
  , FF_ITERATOR    = '@@iterator'
  , KEYS           = 'keys'
  , VALUES         = 'values';

var returnThis = function(){ return this; };

module.exports = function(Base, NAME, Constructor, next, DEFAULT, IS_SET, FORCED){
  $iterCreate(Constructor, NAME, next);
  var getMethod = function(kind){
    if(!BUGGY && kind in proto)return proto[kind];
    switch(kind){
      case KEYS: return function keys(){ return new Constructor(this, kind); };
      case VALUES: return function values(){ return new Constructor(this, kind); };
    } return function entries(){ return new Constructor(this, kind); };
  };
  var TAG        = NAME + ' Iterator'
    , DEF_VALUES = DEFAULT == VALUES
    , VALUES_BUG = false
    , proto      = Base.prototype
    , $native    = proto[ITERATOR] || proto[FF_ITERATOR] || DEFAULT && proto[DEFAULT]
    , $default   = $native || getMethod(DEFAULT)
    , $entries   = DEFAULT ? !DEF_VALUES ? $default : getMethod('entries') : undefined
    , $anyNative = NAME == 'Array' ? proto.entries || $native : $native
    , methods, key, IteratorPrototype;
  // Fix native
  if($anyNative){
    IteratorPrototype = getPrototypeOf($anyNative.call(new Base));
    if(IteratorPrototype !== Object.prototype){
      // Set @@toStringTag to native iterators
      setToStringTag(IteratorPrototype, TAG, true);
      // fix for some old engines
      if(!LIBRARY && !has(IteratorPrototype, ITERATOR))hide(IteratorPrototype, ITERATOR, returnThis);
    }
  }
  // fix Array#{values, @@iterator}.name in V8 / FF
  if(DEF_VALUES && $native && $native.name !== VALUES){
    VALUES_BUG = true;
    $default = function values(){ return $native.call(this); };
  }
  // Define iterator
  if((!LIBRARY || FORCED) && (BUGGY || VALUES_BUG || !proto[ITERATOR])){
    hide(proto, ITERATOR, $default);
  }
  // Plug for library
  Iterators[NAME] = $default;
  Iterators[TAG]  = returnThis;
  if(DEFAULT){
    methods = {
      values:  DEF_VALUES ? $default : getMethod(VALUES),
      keys:    IS_SET     ? $default : getMethod(KEYS),
      entries: $entries
    };
    if(FORCED)for(key in methods){
      if(!(key in proto))redefine(proto, key, methods[key]);
    } else $export($export.P + $export.F * (BUGGY || VALUES_BUG), NAME, methods);
  }
  return methods;
};
},{"./_export":43,"./_has":47,"./_hide":48,"./_iter-create":56,"./_iterators":59,"./_library":61,"./_object-gpo":70,"./_redefine":77,"./_set-to-string-tag":80,"./_wks":93}],58:[function(require,module,exports){
module.exports = function(done, value){
  return {value: value, done: !!done};
};
},{}],59:[function(require,module,exports){
module.exports = {};
},{}],60:[function(require,module,exports){
var getKeys   = require('./_object-keys')
  , toIObject = require('./_to-iobject');
module.exports = function(object, el){
  var O      = toIObject(object)
    , keys   = getKeys(O)
    , length = keys.length
    , index  = 0
    , key;
  while(length > index)if(O[key = keys[index++]] === el)return key;
};
},{"./_object-keys":72,"./_to-iobject":86}],61:[function(require,module,exports){
module.exports = true;
},{}],62:[function(require,module,exports){
var META     = require('./_uid')('meta')
  , isObject = require('./_is-object')
  , has      = require('./_has')
  , setDesc  = require('./_object-dp').f
  , id       = 0;
var isExtensible = Object.isExtensible || function(){
  return true;
};
var FREEZE = !require('./_fails')(function(){
  return isExtensible(Object.preventExtensions({}));
});
var setMeta = function(it){
  setDesc(it, META, {value: {
    i: 'O' + ++id, // object ID
    w: {}          // weak collections IDs
  }});
};
var fastKey = function(it, create){
  // return primitive with prefix
  if(!isObject(it))return typeof it == 'symbol' ? it : (typeof it == 'string' ? 'S' : 'P') + it;
  if(!has(it, META)){
    // can't set metadata to uncaught frozen object
    if(!isExtensible(it))return 'F';
    // not necessary to add metadata
    if(!create)return 'E';
    // add missing metadata
    setMeta(it);
  // return object ID
  } return it[META].i;
};
var getWeak = function(it, create){
  if(!has(it, META)){
    // can't set metadata to uncaught frozen object
    if(!isExtensible(it))return true;
    // not necessary to add metadata
    if(!create)return false;
    // add missing metadata
    setMeta(it);
  // return hash weak collections IDs
  } return it[META].w;
};
// add metadata on freeze-family methods calling
var onFreeze = function(it){
  if(FREEZE && meta.NEED && isExtensible(it) && !has(it, META))setMeta(it);
  return it;
};
var meta = module.exports = {
  KEY:      META,
  NEED:     false,
  fastKey:  fastKey,
  getWeak:  getWeak,
  onFreeze: onFreeze
};
},{"./_fails":44,"./_has":47,"./_is-object":54,"./_object-dp":64,"./_uid":90}],63:[function(require,module,exports){
// 19.1.2.2 / 15.2.3.5 Object.create(O [, Properties])
var anObject    = require('./_an-object')
  , dPs         = require('./_object-dps')
  , enumBugKeys = require('./_enum-bug-keys')
  , IE_PROTO    = require('./_shared-key')('IE_PROTO')
  , Empty       = function(){ /* empty */ }
  , PROTOTYPE   = 'prototype';

// Create object with fake `null` prototype: use iframe Object with cleared prototype
var createDict = function(){
  // Thrash, waste and sodomy: IE GC bug
  var iframe = require('./_dom-create')('iframe')
    , i      = enumBugKeys.length
    , lt     = '<'
    , gt     = '>'
    , iframeDocument;
  iframe.style.display = 'none';
  require('./_html').appendChild(iframe);
  iframe.src = 'javascript:'; // eslint-disable-line no-script-url
  // createDict = iframe.contentWindow.Object;
  // html.removeChild(iframe);
  iframeDocument = iframe.contentWindow.document;
  iframeDocument.open();
  iframeDocument.write(lt + 'script' + gt + 'document.F=Object' + lt + '/script' + gt);
  iframeDocument.close();
  createDict = iframeDocument.F;
  while(i--)delete createDict[PROTOTYPE][enumBugKeys[i]];
  return createDict();
};

module.exports = Object.create || function create(O, Properties){
  var result;
  if(O !== null){
    Empty[PROTOTYPE] = anObject(O);
    result = new Empty;
    Empty[PROTOTYPE] = null;
    // add "__proto__" for Object.getPrototypeOf polyfill
    result[IE_PROTO] = O;
  } else result = createDict();
  return Properties === undefined ? result : dPs(result, Properties);
};

},{"./_an-object":25,"./_dom-create":40,"./_enum-bug-keys":41,"./_html":49,"./_object-dps":65,"./_shared-key":81}],64:[function(require,module,exports){
var anObject       = require('./_an-object')
  , IE8_DOM_DEFINE = require('./_ie8-dom-define')
  , toPrimitive    = require('./_to-primitive')
  , dP             = Object.defineProperty;

exports.f = require('./_descriptors') ? Object.defineProperty : function defineProperty(O, P, Attributes){
  anObject(O);
  P = toPrimitive(P, true);
  anObject(Attributes);
  if(IE8_DOM_DEFINE)try {
    return dP(O, P, Attributes);
  } catch(e){ /* empty */ }
  if('get' in Attributes || 'set' in Attributes)throw TypeError('Accessors not supported!');
  if('value' in Attributes)O[P] = Attributes.value;
  return O;
};
},{"./_an-object":25,"./_descriptors":39,"./_ie8-dom-define":50,"./_to-primitive":89}],65:[function(require,module,exports){
var dP       = require('./_object-dp')
  , anObject = require('./_an-object')
  , getKeys  = require('./_object-keys');

module.exports = require('./_descriptors') ? Object.defineProperties : function defineProperties(O, Properties){
  anObject(O);
  var keys   = getKeys(Properties)
    , length = keys.length
    , i = 0
    , P;
  while(length > i)dP.f(O, P = keys[i++], Properties[P]);
  return O;
};
},{"./_an-object":25,"./_descriptors":39,"./_object-dp":64,"./_object-keys":72}],66:[function(require,module,exports){
var pIE            = require('./_object-pie')
  , createDesc     = require('./_property-desc')
  , toIObject      = require('./_to-iobject')
  , toPrimitive    = require('./_to-primitive')
  , has            = require('./_has')
  , IE8_DOM_DEFINE = require('./_ie8-dom-define')
  , gOPD           = Object.getOwnPropertyDescriptor;

exports.f = require('./_descriptors') ? gOPD : function getOwnPropertyDescriptor(O, P){
  O = toIObject(O);
  P = toPrimitive(P, true);
  if(IE8_DOM_DEFINE)try {
    return gOPD(O, P);
  } catch(e){ /* empty */ }
  if(has(O, P))return createDesc(!pIE.f.call(O, P), O[P]);
};
},{"./_descriptors":39,"./_has":47,"./_ie8-dom-define":50,"./_object-pie":73,"./_property-desc":75,"./_to-iobject":86,"./_to-primitive":89}],67:[function(require,module,exports){
// fallback for IE11 buggy Object.getOwnPropertyNames with iframe and window
var toIObject = require('./_to-iobject')
  , gOPN      = require('./_object-gopn').f
  , toString  = {}.toString;

var windowNames = typeof window == 'object' && window && Object.getOwnPropertyNames
  ? Object.getOwnPropertyNames(window) : [];

var getWindowNames = function(it){
  try {
    return gOPN(it);
  } catch(e){
    return windowNames.slice();
  }
};

module.exports.f = function getOwnPropertyNames(it){
  return windowNames && toString.call(it) == '[object Window]' ? getWindowNames(it) : gOPN(toIObject(it));
};

},{"./_object-gopn":68,"./_to-iobject":86}],68:[function(require,module,exports){
// 19.1.2.7 / 15.2.3.4 Object.getOwnPropertyNames(O)
var $keys      = require('./_object-keys-internal')
  , hiddenKeys = require('./_enum-bug-keys').concat('length', 'prototype');

exports.f = Object.getOwnPropertyNames || function getOwnPropertyNames(O){
  return $keys(O, hiddenKeys);
};
},{"./_enum-bug-keys":41,"./_object-keys-internal":71}],69:[function(require,module,exports){
exports.f = Object.getOwnPropertySymbols;
},{}],70:[function(require,module,exports){
// 19.1.2.9 / 15.2.3.2 Object.getPrototypeOf(O)
var has         = require('./_has')
  , toObject    = require('./_to-object')
  , IE_PROTO    = require('./_shared-key')('IE_PROTO')
  , ObjectProto = Object.prototype;

module.exports = Object.getPrototypeOf || function(O){
  O = toObject(O);
  if(has(O, IE_PROTO))return O[IE_PROTO];
  if(typeof O.constructor == 'function' && O instanceof O.constructor){
    return O.constructor.prototype;
  } return O instanceof Object ? ObjectProto : null;
};
},{"./_has":47,"./_shared-key":81,"./_to-object":88}],71:[function(require,module,exports){
var has          = require('./_has')
  , toIObject    = require('./_to-iobject')
  , arrayIndexOf = require('./_array-includes')(false)
  , IE_PROTO     = require('./_shared-key')('IE_PROTO');

module.exports = function(object, names){
  var O      = toIObject(object)
    , i      = 0
    , result = []
    , key;
  for(key in O)if(key != IE_PROTO)has(O, key) && result.push(key);
  // Don't enum bug & hidden keys
  while(names.length > i)if(has(O, key = names[i++])){
    ~arrayIndexOf(result, key) || result.push(key);
  }
  return result;
};
},{"./_array-includes":27,"./_has":47,"./_shared-key":81,"./_to-iobject":86}],72:[function(require,module,exports){
// 19.1.2.14 / 15.2.3.14 Object.keys(O)
var $keys       = require('./_object-keys-internal')
  , enumBugKeys = require('./_enum-bug-keys');

module.exports = Object.keys || function keys(O){
  return $keys(O, enumBugKeys);
};
},{"./_enum-bug-keys":41,"./_object-keys-internal":71}],73:[function(require,module,exports){
exports.f = {}.propertyIsEnumerable;
},{}],74:[function(require,module,exports){
// most Object methods by ES6 should accept primitives
var $export = require('./_export')
  , core    = require('./_core')
  , fails   = require('./_fails');
module.exports = function(KEY, exec){
  var fn  = (core.Object || {})[KEY] || Object[KEY]
    , exp = {};
  exp[KEY] = exec(fn);
  $export($export.S + $export.F * fails(function(){ fn(1); }), 'Object', exp);
};
},{"./_core":36,"./_export":43,"./_fails":44}],75:[function(require,module,exports){
module.exports = function(bitmap, value){
  return {
    enumerable  : !(bitmap & 1),
    configurable: !(bitmap & 2),
    writable    : !(bitmap & 4),
    value       : value
  };
};
},{}],76:[function(require,module,exports){
var hide = require('./_hide');
module.exports = function(target, src, safe){
  for(var key in src){
    if(safe && target[key])target[key] = src[key];
    else hide(target, key, src[key]);
  } return target;
};
},{"./_hide":48}],77:[function(require,module,exports){
module.exports = require('./_hide');
},{"./_hide":48}],78:[function(require,module,exports){
// Works with __proto__ only. Old v8 can't work with null proto objects.
/* eslint-disable no-proto */
var isObject = require('./_is-object')
  , anObject = require('./_an-object');
var check = function(O, proto){
  anObject(O);
  if(!isObject(proto) && proto !== null)throw TypeError(proto + ": can't set as prototype!");
};
module.exports = {
  set: Object.setPrototypeOf || ('__proto__' in {} ? // eslint-disable-line
    function(test, buggy, set){
      try {
        set = require('./_ctx')(Function.call, require('./_object-gopd').f(Object.prototype, '__proto__').set, 2);
        set(test, []);
        buggy = !(test instanceof Array);
      } catch(e){ buggy = true; }
      return function setPrototypeOf(O, proto){
        check(O, proto);
        if(buggy)O.__proto__ = proto;
        else set(O, proto);
        return O;
      };
    }({}, false) : undefined),
  check: check
};
},{"./_an-object":25,"./_ctx":37,"./_is-object":54,"./_object-gopd":66}],79:[function(require,module,exports){
'use strict';
var global      = require('./_global')
  , core        = require('./_core')
  , dP          = require('./_object-dp')
  , DESCRIPTORS = require('./_descriptors')
  , SPECIES     = require('./_wks')('species');

module.exports = function(KEY){
  var C = typeof core[KEY] == 'function' ? core[KEY] : global[KEY];
  if(DESCRIPTORS && C && !C[SPECIES])dP.f(C, SPECIES, {
    configurable: true,
    get: function(){ return this; }
  });
};
},{"./_core":36,"./_descriptors":39,"./_global":46,"./_object-dp":64,"./_wks":93}],80:[function(require,module,exports){
var def = require('./_object-dp').f
  , has = require('./_has')
  , TAG = require('./_wks')('toStringTag');

module.exports = function(it, tag, stat){
  if(it && !has(it = stat ? it : it.prototype, TAG))def(it, TAG, {configurable: true, value: tag});
};
},{"./_has":47,"./_object-dp":64,"./_wks":93}],81:[function(require,module,exports){
var shared = require('./_shared')('keys')
  , uid    = require('./_uid');
module.exports = function(key){
  return shared[key] || (shared[key] = uid(key));
};
},{"./_shared":82,"./_uid":90}],82:[function(require,module,exports){
var global = require('./_global')
  , SHARED = '__core-js_shared__'
  , store  = global[SHARED] || (global[SHARED] = {});
module.exports = function(key){
  return store[key] || (store[key] = {});
};
},{"./_global":46}],83:[function(require,module,exports){
var toInteger = require('./_to-integer')
  , defined   = require('./_defined');
// true  -> String#at
// false -> String#codePointAt
module.exports = function(TO_STRING){
  return function(that, pos){
    var s = String(defined(that))
      , i = toInteger(pos)
      , l = s.length
      , a, b;
    if(i < 0 || i >= l)return TO_STRING ? '' : undefined;
    a = s.charCodeAt(i);
    return a < 0xd800 || a > 0xdbff || i + 1 === l || (b = s.charCodeAt(i + 1)) < 0xdc00 || b > 0xdfff
      ? TO_STRING ? s.charAt(i) : a
      : TO_STRING ? s.slice(i, i + 2) : (a - 0xd800 << 10) + (b - 0xdc00) + 0x10000;
  };
};
},{"./_defined":38,"./_to-integer":85}],84:[function(require,module,exports){
var toInteger = require('./_to-integer')
  , max       = Math.max
  , min       = Math.min;
module.exports = function(index, length){
  index = toInteger(index);
  return index < 0 ? max(index + length, 0) : min(index, length);
};
},{"./_to-integer":85}],85:[function(require,module,exports){
// 7.1.4 ToInteger
var ceil  = Math.ceil
  , floor = Math.floor;
module.exports = function(it){
  return isNaN(it = +it) ? 0 : (it > 0 ? floor : ceil)(it);
};
},{}],86:[function(require,module,exports){
// to indexed object, toObject with fallback for non-array-like ES3 strings
var IObject = require('./_iobject')
  , defined = require('./_defined');
module.exports = function(it){
  return IObject(defined(it));
};
},{"./_defined":38,"./_iobject":51}],87:[function(require,module,exports){
// 7.1.15 ToLength
var toInteger = require('./_to-integer')
  , min       = Math.min;
module.exports = function(it){
  return it > 0 ? min(toInteger(it), 0x1fffffffffffff) : 0; // pow(2, 53) - 1 == 9007199254740991
};
},{"./_to-integer":85}],88:[function(require,module,exports){
// 7.1.13 ToObject(argument)
var defined = require('./_defined');
module.exports = function(it){
  return Object(defined(it));
};
},{"./_defined":38}],89:[function(require,module,exports){
// 7.1.1 ToPrimitive(input [, PreferredType])
var isObject = require('./_is-object');
// instead of the ES6 spec version, we didn't implement @@toPrimitive case
// and the second argument - flag - preferred type is a string
module.exports = function(it, S){
  if(!isObject(it))return it;
  var fn, val;
  if(S && typeof (fn = it.toString) == 'function' && !isObject(val = fn.call(it)))return val;
  if(typeof (fn = it.valueOf) == 'function' && !isObject(val = fn.call(it)))return val;
  if(!S && typeof (fn = it.toString) == 'function' && !isObject(val = fn.call(it)))return val;
  throw TypeError("Can't convert object to primitive value");
};
},{"./_is-object":54}],90:[function(require,module,exports){
var id = 0
  , px = Math.random();
module.exports = function(key){
  return 'Symbol('.concat(key === undefined ? '' : key, ')_', (++id + px).toString(36));
};
},{}],91:[function(require,module,exports){
var global         = require('./_global')
  , core           = require('./_core')
  , LIBRARY        = require('./_library')
  , wksExt         = require('./_wks-ext')
  , defineProperty = require('./_object-dp').f;
module.exports = function(name){
  var $Symbol = core.Symbol || (core.Symbol = LIBRARY ? {} : global.Symbol || {});
  if(name.charAt(0) != '_' && !(name in $Symbol))defineProperty($Symbol, name, {value: wksExt.f(name)});
};
},{"./_core":36,"./_global":46,"./_library":61,"./_object-dp":64,"./_wks-ext":92}],92:[function(require,module,exports){
exports.f = require('./_wks');
},{"./_wks":93}],93:[function(require,module,exports){
var store      = require('./_shared')('wks')
  , uid        = require('./_uid')
  , Symbol     = require('./_global').Symbol
  , USE_SYMBOL = typeof Symbol == 'function';

var $exports = module.exports = function(name){
  return store[name] || (store[name] =
    USE_SYMBOL && Symbol[name] || (USE_SYMBOL ? Symbol : uid)('Symbol.' + name));
};

$exports.store = store;
},{"./_global":46,"./_shared":82,"./_uid":90}],94:[function(require,module,exports){
var classof   = require('./_classof')
  , ITERATOR  = require('./_wks')('iterator')
  , Iterators = require('./_iterators');
module.exports = require('./_core').getIteratorMethod = function(it){
  if(it != undefined)return it[ITERATOR]
    || it['@@iterator']
    || Iterators[classof(it)];
};
},{"./_classof":31,"./_core":36,"./_iterators":59,"./_wks":93}],95:[function(require,module,exports){
'use strict';
var addToUnscopables = require('./_add-to-unscopables')
  , step             = require('./_iter-step')
  , Iterators        = require('./_iterators')
  , toIObject        = require('./_to-iobject');

// 22.1.3.4 Array.prototype.entries()
// 22.1.3.13 Array.prototype.keys()
// 22.1.3.29 Array.prototype.values()
// 22.1.3.30 Array.prototype[@@iterator]()
module.exports = require('./_iter-define')(Array, 'Array', function(iterated, kind){
  this._t = toIObject(iterated); // target
  this._i = 0;                   // next index
  this._k = kind;                // kind
// 22.1.5.2.1 %ArrayIteratorPrototype%.next()
}, function(){
  var O     = this._t
    , kind  = this._k
    , index = this._i++;
  if(!O || index >= O.length){
    this._t = undefined;
    return step(1);
  }
  if(kind == 'keys'  )return step(0, index);
  if(kind == 'values')return step(0, O[index]);
  return step(0, [index, O[index]]);
}, 'values');

// argumentsList[@@iterator] is %ArrayProto_values% (9.4.4.6, 9.4.4.7)
Iterators.Arguments = Iterators.Array;

addToUnscopables('keys');
addToUnscopables('values');
addToUnscopables('entries');
},{"./_add-to-unscopables":23,"./_iter-define":57,"./_iter-step":58,"./_iterators":59,"./_to-iobject":86}],96:[function(require,module,exports){
'use strict';
var strong = require('./_collection-strong');

// 23.1 Map Objects
module.exports = require('./_collection')('Map', function(get){
  return function Map(){ return get(this, arguments.length > 0 ? arguments[0] : undefined); };
}, {
  // 23.1.3.6 Map.prototype.get(key)
  get: function get(key){
    var entry = strong.getEntry(this, key);
    return entry && entry.v;
  },
  // 23.1.3.9 Map.prototype.set(key, value)
  set: function set(key, value){
    return strong.def(this, key === 0 ? 0 : key, value);
  }
}, strong, true);
},{"./_collection":35,"./_collection-strong":33}],97:[function(require,module,exports){
var $export = require('./_export')
// 19.1.2.2 / 15.2.3.5 Object.create(O [, Properties])
$export($export.S, 'Object', {create: require('./_object-create')});
},{"./_export":43,"./_object-create":63}],98:[function(require,module,exports){
var $export = require('./_export');
// 19.1.2.4 / 15.2.3.6 Object.defineProperty(O, P, Attributes)
$export($export.S + $export.F * !require('./_descriptors'), 'Object', {defineProperty: require('./_object-dp').f});
},{"./_descriptors":39,"./_export":43,"./_object-dp":64}],99:[function(require,module,exports){
// 19.1.2.9 Object.getPrototypeOf(O)
var toObject        = require('./_to-object')
  , $getPrototypeOf = require('./_object-gpo');

require('./_object-sap')('getPrototypeOf', function(){
  return function getPrototypeOf(it){
    return $getPrototypeOf(toObject(it));
  };
});
},{"./_object-gpo":70,"./_object-sap":74,"./_to-object":88}],100:[function(require,module,exports){
// 19.1.3.19 Object.setPrototypeOf(O, proto)
var $export = require('./_export');
$export($export.S, 'Object', {setPrototypeOf: require('./_set-proto').set});
},{"./_export":43,"./_set-proto":78}],101:[function(require,module,exports){

},{}],102:[function(require,module,exports){
'use strict';
var $at  = require('./_string-at')(true);

// 21.1.3.27 String.prototype[@@iterator]()
require('./_iter-define')(String, 'String', function(iterated){
  this._t = String(iterated); // target
  this._i = 0;                // next index
// 21.1.5.2.1 %StringIteratorPrototype%.next()
}, function(){
  var O     = this._t
    , index = this._i
    , point;
  if(index >= O.length)return {value: undefined, done: true};
  point = $at(O, index);
  this._i += point.length;
  return {value: point, done: false};
});
},{"./_iter-define":57,"./_string-at":83}],103:[function(require,module,exports){
'use strict';
// ECMAScript 6 symbols shim
var global         = require('./_global')
  , has            = require('./_has')
  , DESCRIPTORS    = require('./_descriptors')
  , $export        = require('./_export')
  , redefine       = require('./_redefine')
  , META           = require('./_meta').KEY
  , $fails         = require('./_fails')
  , shared         = require('./_shared')
  , setToStringTag = require('./_set-to-string-tag')
  , uid            = require('./_uid')
  , wks            = require('./_wks')
  , wksExt         = require('./_wks-ext')
  , wksDefine      = require('./_wks-define')
  , keyOf          = require('./_keyof')
  , enumKeys       = require('./_enum-keys')
  , isArray        = require('./_is-array')
  , anObject       = require('./_an-object')
  , toIObject      = require('./_to-iobject')
  , toPrimitive    = require('./_to-primitive')
  , createDesc     = require('./_property-desc')
  , _create        = require('./_object-create')
  , gOPNExt        = require('./_object-gopn-ext')
  , $GOPD          = require('./_object-gopd')
  , $DP            = require('./_object-dp')
  , $keys          = require('./_object-keys')
  , gOPD           = $GOPD.f
  , dP             = $DP.f
  , gOPN           = gOPNExt.f
  , $Symbol        = global.Symbol
  , $JSON          = global.JSON
  , _stringify     = $JSON && $JSON.stringify
  , PROTOTYPE      = 'prototype'
  , HIDDEN         = wks('_hidden')
  , TO_PRIMITIVE   = wks('toPrimitive')
  , isEnum         = {}.propertyIsEnumerable
  , SymbolRegistry = shared('symbol-registry')
  , AllSymbols     = shared('symbols')
  , OPSymbols      = shared('op-symbols')
  , ObjectProto    = Object[PROTOTYPE]
  , USE_NATIVE     = typeof $Symbol == 'function'
  , QObject        = global.QObject;
// Don't use setters in Qt Script, https://github.com/zloirock/core-js/issues/173
var setter = !QObject || !QObject[PROTOTYPE] || !QObject[PROTOTYPE].findChild;

// fallback for old Android, https://code.google.com/p/v8/issues/detail?id=687
var setSymbolDesc = DESCRIPTORS && $fails(function(){
  return _create(dP({}, 'a', {
    get: function(){ return dP(this, 'a', {value: 7}).a; }
  })).a != 7;
}) ? function(it, key, D){
  var protoDesc = gOPD(ObjectProto, key);
  if(protoDesc)delete ObjectProto[key];
  dP(it, key, D);
  if(protoDesc && it !== ObjectProto)dP(ObjectProto, key, protoDesc);
} : dP;

var wrap = function(tag){
  var sym = AllSymbols[tag] = _create($Symbol[PROTOTYPE]);
  sym._k = tag;
  return sym;
};

var isSymbol = USE_NATIVE && typeof $Symbol.iterator == 'symbol' ? function(it){
  return typeof it == 'symbol';
} : function(it){
  return it instanceof $Symbol;
};

var $defineProperty = function defineProperty(it, key, D){
  if(it === ObjectProto)$defineProperty(OPSymbols, key, D);
  anObject(it);
  key = toPrimitive(key, true);
  anObject(D);
  if(has(AllSymbols, key)){
    if(!D.enumerable){
      if(!has(it, HIDDEN))dP(it, HIDDEN, createDesc(1, {}));
      it[HIDDEN][key] = true;
    } else {
      if(has(it, HIDDEN) && it[HIDDEN][key])it[HIDDEN][key] = false;
      D = _create(D, {enumerable: createDesc(0, false)});
    } return setSymbolDesc(it, key, D);
  } return dP(it, key, D);
};
var $defineProperties = function defineProperties(it, P){
  anObject(it);
  var keys = enumKeys(P = toIObject(P))
    , i    = 0
    , l = keys.length
    , key;
  while(l > i)$defineProperty(it, key = keys[i++], P[key]);
  return it;
};
var $create = function create(it, P){
  return P === undefined ? _create(it) : $defineProperties(_create(it), P);
};
var $propertyIsEnumerable = function propertyIsEnumerable(key){
  var E = isEnum.call(this, key = toPrimitive(key, true));
  if(this === ObjectProto && has(AllSymbols, key) && !has(OPSymbols, key))return false;
  return E || !has(this, key) || !has(AllSymbols, key) || has(this, HIDDEN) && this[HIDDEN][key] ? E : true;
};
var $getOwnPropertyDescriptor = function getOwnPropertyDescriptor(it, key){
  it  = toIObject(it);
  key = toPrimitive(key, true);
  if(it === ObjectProto && has(AllSymbols, key) && !has(OPSymbols, key))return;
  var D = gOPD(it, key);
  if(D && has(AllSymbols, key) && !(has(it, HIDDEN) && it[HIDDEN][key]))D.enumerable = true;
  return D;
};
var $getOwnPropertyNames = function getOwnPropertyNames(it){
  var names  = gOPN(toIObject(it))
    , result = []
    , i      = 0
    , key;
  while(names.length > i){
    if(!has(AllSymbols, key = names[i++]) && key != HIDDEN && key != META)result.push(key);
  } return result;
};
var $getOwnPropertySymbols = function getOwnPropertySymbols(it){
  var IS_OP  = it === ObjectProto
    , names  = gOPN(IS_OP ? OPSymbols : toIObject(it))
    , result = []
    , i      = 0
    , key;
  while(names.length > i){
    if(has(AllSymbols, key = names[i++]) && (IS_OP ? has(ObjectProto, key) : true))result.push(AllSymbols[key]);
  } return result;
};

// 19.4.1.1 Symbol([description])
if(!USE_NATIVE){
  $Symbol = function Symbol(){
    if(this instanceof $Symbol)throw TypeError('Symbol is not a constructor!');
    var tag = uid(arguments.length > 0 ? arguments[0] : undefined);
    var $set = function(value){
      if(this === ObjectProto)$set.call(OPSymbols, value);
      if(has(this, HIDDEN) && has(this[HIDDEN], tag))this[HIDDEN][tag] = false;
      setSymbolDesc(this, tag, createDesc(1, value));
    };
    if(DESCRIPTORS && setter)setSymbolDesc(ObjectProto, tag, {configurable: true, set: $set});
    return wrap(tag);
  };
  redefine($Symbol[PROTOTYPE], 'toString', function toString(){
    return this._k;
  });

  $GOPD.f = $getOwnPropertyDescriptor;
  $DP.f   = $defineProperty;
  require('./_object-gopn').f = gOPNExt.f = $getOwnPropertyNames;
  require('./_object-pie').f  = $propertyIsEnumerable;
  require('./_object-gops').f = $getOwnPropertySymbols;

  if(DESCRIPTORS && !require('./_library')){
    redefine(ObjectProto, 'propertyIsEnumerable', $propertyIsEnumerable, true);
  }

  wksExt.f = function(name){
    return wrap(wks(name));
  }
}

$export($export.G + $export.W + $export.F * !USE_NATIVE, {Symbol: $Symbol});

for(var symbols = (
  // 19.4.2.2, 19.4.2.3, 19.4.2.4, 19.4.2.6, 19.4.2.8, 19.4.2.9, 19.4.2.10, 19.4.2.11, 19.4.2.12, 19.4.2.13, 19.4.2.14
  'hasInstance,isConcatSpreadable,iterator,match,replace,search,species,split,toPrimitive,toStringTag,unscopables'
).split(','), i = 0; symbols.length > i; )wks(symbols[i++]);

for(var symbols = $keys(wks.store), i = 0; symbols.length > i; )wksDefine(symbols[i++]);

$export($export.S + $export.F * !USE_NATIVE, 'Symbol', {
  // 19.4.2.1 Symbol.for(key)
  'for': function(key){
    return has(SymbolRegistry, key += '')
      ? SymbolRegistry[key]
      : SymbolRegistry[key] = $Symbol(key);
  },
  // 19.4.2.5 Symbol.keyFor(sym)
  keyFor: function keyFor(key){
    if(isSymbol(key))return keyOf(SymbolRegistry, key);
    throw TypeError(key + ' is not a symbol!');
  },
  useSetter: function(){ setter = true; },
  useSimple: function(){ setter = false; }
});

$export($export.S + $export.F * !USE_NATIVE, 'Object', {
  // 19.1.2.2 Object.create(O [, Properties])
  create: $create,
  // 19.1.2.4 Object.defineProperty(O, P, Attributes)
  defineProperty: $defineProperty,
  // 19.1.2.3 Object.defineProperties(O, Properties)
  defineProperties: $defineProperties,
  // 19.1.2.6 Object.getOwnPropertyDescriptor(O, P)
  getOwnPropertyDescriptor: $getOwnPropertyDescriptor,
  // 19.1.2.7 Object.getOwnPropertyNames(O)
  getOwnPropertyNames: $getOwnPropertyNames,
  // 19.1.2.8 Object.getOwnPropertySymbols(O)
  getOwnPropertySymbols: $getOwnPropertySymbols
});

// 24.3.2 JSON.stringify(value [, replacer [, space]])
$JSON && $export($export.S + $export.F * (!USE_NATIVE || $fails(function(){
  var S = $Symbol();
  // MS Edge converts symbol values to JSON as {}
  // WebKit converts symbol values to JSON as null
  // V8 throws on boxed symbols
  return _stringify([S]) != '[null]' || _stringify({a: S}) != '{}' || _stringify(Object(S)) != '{}';
})), 'JSON', {
  stringify: function stringify(it){
    if(it === undefined || isSymbol(it))return; // IE8 returns string on undefined
    var args = [it]
      , i    = 1
      , replacer, $replacer;
    while(arguments.length > i)args.push(arguments[i++]);
    replacer = args[1];
    if(typeof replacer == 'function')$replacer = replacer;
    if($replacer || !isArray(replacer))replacer = function(key, value){
      if($replacer)value = $replacer.call(this, key, value);
      if(!isSymbol(value))return value;
    };
    args[1] = replacer;
    return _stringify.apply($JSON, args);
  }
});

// 19.4.3.4 Symbol.prototype[@@toPrimitive](hint)
$Symbol[PROTOTYPE][TO_PRIMITIVE] || require('./_hide')($Symbol[PROTOTYPE], TO_PRIMITIVE, $Symbol[PROTOTYPE].valueOf);
// 19.4.3.5 Symbol.prototype[@@toStringTag]
setToStringTag($Symbol, 'Symbol');
// 20.2.1.9 Math[@@toStringTag]
setToStringTag(Math, 'Math', true);
// 24.3.3 JSON[@@toStringTag]
setToStringTag(global.JSON, 'JSON', true);
},{"./_an-object":25,"./_descriptors":39,"./_enum-keys":42,"./_export":43,"./_fails":44,"./_global":46,"./_has":47,"./_hide":48,"./_is-array":53,"./_keyof":60,"./_library":61,"./_meta":62,"./_object-create":63,"./_object-dp":64,"./_object-gopd":66,"./_object-gopn":68,"./_object-gopn-ext":67,"./_object-gops":69,"./_object-keys":72,"./_object-pie":73,"./_property-desc":75,"./_redefine":77,"./_set-to-string-tag":80,"./_shared":82,"./_to-iobject":86,"./_to-primitive":89,"./_uid":90,"./_wks":93,"./_wks-define":91,"./_wks-ext":92}],104:[function(require,module,exports){
// https://github.com/DavidBruant/Map-Set.prototype.toJSON
var $export  = require('./_export');

$export($export.P + $export.R, 'Map', {toJSON: require('./_collection-to-json')('Map')});
},{"./_collection-to-json":34,"./_export":43}],105:[function(require,module,exports){
require('./_wks-define')('asyncIterator');
},{"./_wks-define":91}],106:[function(require,module,exports){
require('./_wks-define')('observable');
},{"./_wks-define":91}],107:[function(require,module,exports){
require('./es6.array.iterator');
var global        = require('./_global')
  , hide          = require('./_hide')
  , Iterators     = require('./_iterators')
  , TO_STRING_TAG = require('./_wks')('toStringTag');

for(var collections = ['NodeList', 'DOMTokenList', 'MediaList', 'StyleSheetList', 'CSSRuleList'], i = 0; i < 5; i++){
  var NAME       = collections[i]
    , Collection = global[NAME]
    , proto      = Collection && Collection.prototype;
  if(proto && !proto[TO_STRING_TAG])hide(proto, TO_STRING_TAG, NAME);
  Iterators[NAME] = Iterators.Array;
}
},{"./_global":46,"./_hide":48,"./_iterators":59,"./_wks":93,"./es6.array.iterator":95}],108:[function(require,module,exports){
arguments[4][101][0].apply(exports,arguments)
},{"dup":101}],109:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');

/**
 * @param {{
 *  name: string,
 *  value?: *,
 *  children?: Array.<string>,
 *  attributes?: Array.<string, *>,
 *  ns?: string
 * }} config
 */
function toXMLString(config) {
	var name = config.name;
	var string = '<' + name;
	var content = '';
	var attr = void 0;
	var i = void 0,
	    l = void 0;

	if (config.ns) {
		string = util.xmlPrefix + string + ' xmlns="' + util.schemas[config.ns] + '"';
	}
	if (config.attributes) {
		for (i = 0, l = config.attributes.length; i < l; i++) {
			attr = config.attributes[i];

			string += ' ' + attr[0] + '="' + _.escape(attr[1]) + '"';
		}
	}
	if (!_.isUndefined(config.value)) {
		content += _.escape(config.value);
	}
	if (config.children) {
		content += config.children.join('');
	}

	if (content) {
		string += '>' + content + '</' + name + '>';
	} else {
		string += '/>';
	}

	return string;
}

module.exports = toXMLString;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./util":134}],110:[function(require,module,exports){
'use strict';

var _create = require('babel-runtime/core-js/object/create');

var _create2 = _interopRequireDefault(_create);

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var Paths = require('./paths');

var Images = function (_Paths) {
	(0, _inherits3.default)(Images, _Paths);

	function Images() {
		(0, _classCallCheck3.default)(this, Images);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Images.__proto__ || (0, _getPrototypeOf2.default)(Images)).call(this));

		_this._images = (0, _create2.default)(null);
		_this._imageByNames = (0, _create2.default)(null);
		_this._extensions = (0, _create2.default)(null);
		return _this;
	}

	(0, _createClass3.default)(Images, [{
		key: 'addImage',
		value: function addImage(data, type, name) {
			var image = this._images[data];

			if (!image) {
				type = type || '';

				var id = this.uniqueIdForSpace('image');
				var path = 'xl/media/image' + id + '.' + type;
				var contentType = void 0;

				name = name || 'excelWriter' + id;
				switch (type.toLowerCase()) {
					case 'jpeg':
					case 'jpg':
						contentType = 'image/jpeg';
						break;
					case 'png':
						contentType = 'image/png';
						break;
					case 'gif':
						contentType = 'image/gif';
						break;
					default:
						contentType = null;
						break;
				}

				image = {
					objectId: 'image' + id,
					data: data,
					name: name,
					contentType: contentType,
					extension: type,
					path: path
				};
				this.addPath(image, '/' + path);
				this._images[data] = image;
				this._imageByNames[name] = image;
				this._extensions[type] = contentType;
			} else if (name && !this._imageByNames[name]) {
				image.name = name;
				this._imageByNames[name] = image;
			}
			return image.name;
		}
	}, {
		key: 'getImage',
		value: function getImage(name) {
			return this._imageByNames[name];
		}
	}, {
		key: 'getImages',
		value: function getImages() {
			return this._images;
		}
	}, {
		key: 'removeImages',
		value: function removeImages() {
			this._images = null;
			this._imageByNames = null;
		}
	}, {
		key: 'getExtensions',
		value: function getExtensions() {
			return this._extensions;
		}
	}]);
	return Images;
}(Paths);

module.exports = Images;

},{"./paths":112,"babel-runtime/core-js/object/create":3,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],111:[function(require,module,exports){
'use strict';

var _create = require('babel-runtime/core-js/object/create');

var _create2 = _interopRequireDefault(_create);

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var Images = require('./images');
var SharedStrings = require('./sharedStrings');
var Styles = require('../styles');

var Common = function (_Images) {
	(0, _inherits3.default)(Common, _Images);

	function Common() {
		(0, _classCallCheck3.default)(this, Common);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Common.__proto__ || (0, _getPrototypeOf2.default)(Common)).call(this));

		_this.idSpaces = (0, _create2.default)(null);

		_this.sharedStrings = new SharedStrings(_this);
		_this.addPath(_this.sharedStrings, 'sharedStrings.xml');

		_this.styles = new Styles(_this);
		_this.addPath(_this.styles, 'styles.xml');

		_this.worksheets = [];
		_this.tables = [];
		_this.drawings = [];
		return _this;
	}

	(0, _createClass3.default)(Common, [{
		key: 'uniqueId',
		value: function uniqueId(space) {
			return space + this.uniqueIdForSpace(space);
		}
	}, {
		key: 'uniqueIdForSpace',
		value: function uniqueIdForSpace(space) {
			if (!this.idSpaces[space]) {
				this.idSpaces[space] = 1;
			}
			return this.idSpaces[space]++;
		}
	}, {
		key: 'addString',
		value: function addString(string) {
			return this.sharedStrings.add(string);
		}
	}, {
		key: 'addWorksheet',
		value: function addWorksheet(worksheet) {
			var index = this.worksheets.length + 1;
			var path = 'worksheets/sheet' + index + '.xml';
			var relationsPath = 'xl/worksheets/_rels/sheet' + index + '.xml.rels';

			worksheet.path = 'xl/' + path;
			worksheet.relationsPath = relationsPath;
			this.worksheets.push(worksheet);
			this.addPath(worksheet, path);
		}
	}, {
		key: 'getNewWorksheetDefaultName',
		value: function getNewWorksheetDefaultName() {
			return 'Sheet ' + (this.worksheets.length + 1);
		}
	}, {
		key: 'setActiveWorksheet',
		value: function setActiveWorksheet(worksheet) {
			this.activeWorksheet = worksheet;
		}
	}, {
		key: 'addTable',
		value: function addTable(table) {
			var index = this.tables.length + 1;
			var path = 'xl/tables/table' + index + '.xml';

			table.path = path;
			this.tables.push(table);
			this.addPath(table, '/' + path);
		}
	}, {
		key: 'addDrawings',
		value: function addDrawings(drawings) {
			var index = this.drawings.length + 1;
			var path = 'xl/drawings/drawing' + index + '.xml';
			var relationsPath = 'xl/drawings/_rels/drawing' + index + '.xml.rels';

			drawings.path = path;
			drawings.relationsPath = relationsPath;
			this.drawings.push(drawings);
			this.addPath(drawings, '/' + path);
		}
	}]);
	return Common;
}(Images);

module.exports = Common;

},{"../styles":126,"./images":110,"./sharedStrings":113,"babel-runtime/core-js/object/create":3,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],112:[function(require,module,exports){
'use strict';

var _map = require('babel-runtime/core-js/map');

var _map2 = _interopRequireDefault(_map);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var Paths = function () {
	function Paths() {
		(0, _classCallCheck3.default)(this, Paths);

		this._paths = new _map2.default();
	}

	(0, _createClass3.default)(Paths, [{
		key: 'addPath',
		value: function addPath(object, path) {
			this._paths.set(object.objectId, path);
		}
	}, {
		key: 'getPath',
		value: function getPath(object) {
			return this._paths.get(object.objectId);
		}
	}]);
	return Paths;
}();

module.exports = Paths;

},{"babel-runtime/core-js/map":2,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],113:[function(require,module,exports){
(function (global){
/*eslint global-require: "off"*/

'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

var _map = require('babel-runtime/core-js/map');

var _map2 = _interopRequireDefault(_map);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var Readable = require('stream').Readable || null;
var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');

var spaceRE = /^\s|\s$/;

var SharedStrings = function () {
	function SharedStrings(common) {
		(0, _classCallCheck3.default)(this, SharedStrings);

		this.objectId = common.uniqueId('SharedStrings');
		this._strings = new _map2.default();
		this._stringArray = [];
	}

	(0, _createClass3.default)(SharedStrings, [{
		key: 'add',
		value: function add(string) {
			if (!this._strings.has(string)) {
				var stringId = this._stringArray.length;

				this._strings.set(string, stringId);
				this._stringArray[stringId] = string;
			}

			return this._strings.get(string);
		}
	}, {
		key: 'isEmpty',
		value: function isEmpty() {
			return this._stringArray.length === 0;
		}
	}, {
		key: 'save',
		value: function save(canStream) {
			this._strings = null;

			if (canStream) {
				return new SharedStringsStream({
					strings: this._stringArray
				});
			}
			return getXMLBegin(this._stringArray.length) + _.map(this._stringArray, function (string) {
				string = _.escape(string);

				if (spaceRE.test(string)) {
					return '<si><t xml:space="preserve">' + string + '</t></si>';
				}
				return '<si><t>' + string + '</t></si>';
			}).join('') + getXMLEnd();
		}
	}]);
	return SharedStrings;
}();

var SharedStringsStream = function (_Readable) {
	(0, _inherits3.default)(SharedStringsStream, _Readable);

	function SharedStringsStream(options) {
		(0, _classCallCheck3.default)(this, SharedStringsStream);

		var _this = (0, _possibleConstructorReturn3.default)(this, (SharedStringsStream.__proto__ || (0, _getPrototypeOf2.default)(SharedStringsStream)).call(this, options));

		_this.strings = options.strings;
		_this.status = 0;
		return _this;
	}

	(0, _createClass3.default)(SharedStringsStream, [{
		key: '_read',
		value: function _read(size) {
			var stop = false;

			if (this.status === 0) {
				stop = !this.push(getXMLBegin(this.strings.length));

				this.status = 1;
				this.index = 0;
				this.len = this.strings.length;
			}

			if (this.status === 1) {
				var s = '';
				var str = void 0;

				while (this.index < this.len && !stop) {
					while (this.index < this.len && s.length < size) {
						str = _.escape(this.strings[this.index]);

						if (spaceRE.test(str)) {
							s += '<si><t xml:space="preserve">' + str + '</t></si>';
						} else {
							s += '<si><t>' + str + '</t></si>';
						}
						this.strings[this.index] = null;
						this.index++;
					}
					stop = !this.push(s);
					s = '';
				}

				if (this.index === this.len) {
					this.status = 2;
				}
			}

			if (this.status === 2) {
				this.push(getXMLEnd());
				this.push(null);
			}
		}
	}]);
	return SharedStringsStream;
}(Readable);

function getXMLBegin(length) {
	return util.xmlPrefix + '<sst xmlns="' + util.schemas.spreadsheetml + '" count="' + length + '" uniqueCount="' + length + '">';
}

function getXMLEnd() {
	return '</sst>';
}

module.exports = SharedStrings;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../util":134,"babel-runtime/core-js/map":2,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12,"stream":108}],114:[function(require,module,exports){
(function (global){
'use strict';

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

var Anchor = function () {
	function Anchor(config) {
		(0, _classCallCheck3.default)(this, Anchor);

		var coord = void 0;

		if (_.isObject(config)) {
			if (_.has(config, 'cell')) {
				if (_.isObject(config.cell)) {
					coord = { x: config.cell.c || 1, y: config.cell.r || 1 };
				} else {
					coord = util.letterToPosition(config.cell || '');
				}
			} else {
				coord = { x: config.c || 1, y: config.r || 1 };
			}
		} else {
			coord = util.letterToPosition(config || '');
			config = {};
		}
		var x = coord.x - 1;
		var y = coord.y - 1;

		this.from = {
			x: x,
			y: y,
			xOff: util.pixelsToEMUs(config.left || 0),
			yOff: util.pixelsToEMUs(config.top || 0)
		};
		this.to = {
			x: x + (config.cols || 1),
			y: y + (config.rows || 1),
			xOff: util.pixelsToEMUs(-config.right || 0),
			yOff: util.pixelsToEMUs(-config.bottom || 0)
		};
	}

	(0, _createClass3.default)(Anchor, [{
		key: 'saveWithContent',
		value: function saveWithContent(content) {
			return toXMLString({
				name: 'xdr:twoCellAnchor',
				children: [toXMLString({
					name: 'xdr:from',
					children: [toXMLString({
						name: 'xdr:col',
						value: this.from.x
					}), toXMLString({
						name: 'xdr:colOff',
						value: this.from.xOff
					}), toXMLString({
						name: 'xdr:row',
						value: this.from.y
					}), toXMLString({
						name: 'xdr:rowOff',
						value: this.from.yOff
					})]
				}), toXMLString({
					name: 'xdr:to',
					children: [toXMLString({
						name: 'xdr:col',
						value: this.to.x
					}), toXMLString({
						name: 'xdr:colOff',
						value: this.to.xOff
					}), toXMLString({
						name: 'xdr:row',
						value: this.to.y
					}), toXMLString({
						name: 'xdr:rowOff',
						value: this.to.yOff
					})]
				}), content, toXMLString({
					name: 'xdr:clientData'
				})]
			});
		}
	}]);
	return Anchor;
}();

module.exports = Anchor;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"../util":134,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],115:[function(require,module,exports){
'use strict';

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var util = require('../util');
var toXMLString = require('../XMLString');

var AnchorAbsolute = function () {
	function AnchorAbsolute(config) {
		(0, _classCallCheck3.default)(this, AnchorAbsolute);

		config = config || {};

		this.x = util.pixelsToEMUs(config.left || 0);
		this.y = util.pixelsToEMUs(config.top || 0);
		this.width = util.pixelsToEMUs(config.width || 0);
		this.height = util.pixelsToEMUs(config.height || 0);
	}

	(0, _createClass3.default)(AnchorAbsolute, [{
		key: 'saveWithContent',
		value: function saveWithContent(content) {
			return toXMLString({
				name: 'xdr:absoluteAnchor',
				children: [toXMLString({
					name: 'xdr:pos',
					attributes: [['x', this.x], ['y', this.y]]
				}), toXMLString({
					name: 'xdr:ext',
					attributes: [['cx', this.width], ['cy', this.height]]
				}), content, toXMLString({
					name: 'xdr:clientData'
				})]
			});
		}
	}]);
	return AnchorAbsolute;
}();

module.exports = AnchorAbsolute;

},{"../XMLString":109,"../util":134,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],116:[function(require,module,exports){
(function (global){
'use strict';

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var toXMLString = require('../XMLString');

var AnchorOneCell = function () {
	function AnchorOneCell(config) {
		(0, _classCallCheck3.default)(this, AnchorOneCell);

		var coord = void 0;

		if (_.isObject(config)) {
			if (_.has(config, 'cell')) {
				if (_.isObject(config.cell)) {
					coord = { x: config.cell.c || 1, y: config.cell.r || 1 };
				} else {
					coord = util.letterToPosition(config.cell || '');
				}
			} else {
				coord = { x: config.c || 1, y: config.r || 1 };
			}
		} else {
			coord = util.letterToPosition(config || '');
			config = {};
		}

		this.x = coord.x - 1;
		this.y = coord.y - 1;
		this.xOff = util.pixelsToEMUs(config.left || 0);
		this.yOff = util.pixelsToEMUs(config.top || 0);
		this.width = util.pixelsToEMUs(config.width || 0);
		this.height = util.pixelsToEMUs(config.height || 0);
	}

	(0, _createClass3.default)(AnchorOneCell, [{
		key: 'saveWithContent',
		value: function saveWithContent(content) {
			return toXMLString({
				name: 'xdr:oneCellAnchor',
				children: [toXMLString({
					name: 'xdr:from',
					children: [toXMLString({
						name: 'xdr:col',
						value: this.x
					}), toXMLString({
						name: 'xdr:colOff',
						value: this.xOff
					}), toXMLString({
						name: 'xdr:row',
						value: this.y
					}), toXMLString({
						name: 'xdr:rowOff',
						value: this.yOff
					})]
				}), toXMLString({
					name: 'xdr:ext',
					attributes: [['cx', this.width], ['cy', this.height]]
				}), content, toXMLString({
					name: 'xdr:clientData'
				})]
			});
		}
	}]);
	return AnchorOneCell;
}();

module.exports = AnchorOneCell;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"../util":134,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],117:[function(require,module,exports){
'use strict';

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var Relations = require('../relations');
var util = require('../util');
var toXMLString = require('../XMLString');
var Picture = require('./picture');

var Drawings = function () {
	function Drawings(common) {
		(0, _classCallCheck3.default)(this, Drawings);

		this.common = common;

		this.objectId = this.common.uniqueId('Drawings');
		this.drawings = [];
		this.relations = new Relations(common);
	}

	(0, _createClass3.default)(Drawings, [{
		key: 'addImage',
		value: function addImage(name, config, anchorType) {
			var image = this.common.getImage(name);
			var imageRelationId = this.relations.addRelation(image, 'image');
			var picture = new Picture(this.common, {
				image: image,
				imageRelationId: imageRelationId,
				config: config,
				anchorType: anchorType
			});

			this.drawings.push(picture);
		}
	}, {
		key: 'save',
		value: function save() {
			var attributes = [['xmlns:a', util.schemas.drawing], ['xmlns:r', util.schemas.relationships], ['xmlns:xdr', util.schemas.spreadsheetDrawing]];
			var children = this.drawings.map(function (picture) {
				return picture.save();
			});

			return toXMLString({
				name: 'xdr:wsDr',
				ns: 'spreadsheetDrawing',
				attributes: attributes,
				children: children
			});
		}
	}]);
	return Drawings;
}();

module.exports = Drawings;

},{"../XMLString":109,"../relations":120,"../util":134,"./picture":118,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],118:[function(require,module,exports){
'use strict';

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var util = require('../util');
var toXMLString = require('../XMLString');
var Anchor = require('./anchor');
var AnchorOneCell = require('./anchorOneCell');
var AnchorAbsolute = require('./anchorAbsolute');

var Picture = function () {
	function Picture(common, config) {
		(0, _classCallCheck3.default)(this, Picture);

		this.pictureId = common.uniqueIdForSpace('Picture');
		this.image = config.image;
		this.imageRelationId = config.imageRelationId;
		this.createAnchor(config.anchorType, config.config);
	}
	/**
  *
  * @param {String} type Can be 'anchor', 'oneCell' or 'absolute'.
  * @param {Object} config Shorthand - pass the created anchor coords that can normally be used to construct it.
  */


	(0, _createClass3.default)(Picture, [{
		key: 'createAnchor',
		value: function createAnchor(type, config) {
			switch (type) {
				case 'anchor':
					this.anchor = new Anchor(config);
					break;
				case 'oneCell':
					this.anchor = new AnchorOneCell(config);
					break;
				case 'absolute':
					this.anchor = new AnchorAbsolute(config);
					break;
			}
		}
	}, {
		key: 'save',
		value: function save() {
			var picture = toXMLString({
				name: 'xdr:pic',
				children: [toXMLString({
					name: 'xdr:nvPicPr',
					children: [toXMLString({
						name: 'xdr:cNvPr',
						attributes: [['id', this.pictureId], ['name', this.image.name]]
					}), toXMLString({
						name: 'xdr:cNvPicPr',
						children: [toXMLString({
							name: 'a:picLocks',
							attributes: [['noChangeAspect', '1'], ['noChangeArrowheads', '1']]
						})]
					})]
				}), toXMLString({
					name: 'xdr:blipFill',
					children: [toXMLString({
						name: 'a:blip',
						attributes: [['xmlns:r', util.schemas.relationships], ['r:embed', this.imageRelationId]]
					}), toXMLString({
						name: 'a:srcRect'
					}), toXMLString({
						name: 'a:stretch',
						children: [toXMLString({
							name: 'a:fillRect'
						})]
					})]
				}), toXMLString({
					name: 'xdr:spPr',
					attributes: [['bwMode', 'auto']],
					children: [toXMLString({
						name: 'a:xfrm'
					}), toXMLString({
						name: 'a:prstGeom',
						attributes: [['prst', 'rect']]
					})]
				})]
			});

			return this.anchor.saveWithContent(picture);
		}
	}]);
	return Picture;
}();

module.exports = Picture;

},{"../XMLString":109,"../util":134,"./anchor":114,"./anchorAbsolute":115,"./anchorOneCell":116,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],119:[function(require,module,exports){
(function (global){
'use strict';

var Readable = require('stream').Readable;
var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var JSZip = typeof window !== "undefined" ? window['JSZip'] : typeof global !== "undefined" ? global['JSZip'] : null;
var Workbook = require('./workbook');

module.exports = {
	createWorkbook: function createWorkbook() {
		return new Workbook();
	},

	/**
  * Turns a workbook into a downloadable file.
  * @param {Workbook} workbook - The workbook that is being converted
  * @param {Object?} options - options to modify how the zip is created. See http://stuk.github.io/jszip/#doc_generate_options
  */
	save: function save(workbook, options) {
		var zip = new JSZip();
		var canStream = !!Readable;

		workbook._generateFiles(zip, canStream);
		return zip.generateAsync(_.defaults(options, {
			compression: 'DEFLATE',
			mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
			type: 'base64'
		}));
	},

	saveAsNodeStream: function saveAsNodeStream(workbook, options) {
		var zip = new JSZip();
		var canStream = !!Readable;

		workbook._generateFiles(zip, canStream);
		return zip.generateNodeStream(_.defaults(options, {
			compression: 'DEFLATE'
		}));
	}
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./workbook":135,"stream":108}],120:[function(require,module,exports){
(function (global){
'use strict';

var _create = require('babel-runtime/core-js/object/create');

var _create2 = _interopRequireDefault(_create);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');
var toXMLString = require('./XMLString');

var RelationshipManager = function () {
	function RelationshipManager(common) {
		(0, _classCallCheck3.default)(this, RelationshipManager);

		this.common = common;

		this.relations = (0, _create2.default)(null);
		this.lastId = 1;
	}

	(0, _createClass3.default)(RelationshipManager, [{
		key: 'addRelation',
		value: function addRelation(object, type) {
			var relation = this.relations[object.objectId];

			if (relation) {
				return relation.relationId;
			}

			var relationId = 'rId' + this.lastId++;
			this.relations[object.objectId] = {
				relationId: relationId,
				schema: util.schemas[type],
				object: object
			};
			return relationId;
		}
	}, {
		key: 'getRelationshipId',
		value: function getRelationshipId(object) {
			var relation = this.relations[object.objectId];

			return relation ? relation.relationId : null;
		}
	}, {
		key: 'save',
		value: function save() {
			var common = this.common;
			var children = _.map(this.relations, function (relation) {
				var attributes = [['Id', relation.relationId], ['Type', relation.schema], ['Target', relation.object.target || common.getPath(relation.object)]];

				if (relation.object.targetMode) {
					attributes.push(['TargetMode', relation.object.targetMode]);
				}

				return toXMLString({
					name: 'Relationship',
					attributes: attributes
				});
			});

			return toXMLString({
				name: 'Relationships',
				ns: 'relationshipPackage',
				children: children
			});
		}
	}]);
	return RelationshipManager;
}();

module.exports = RelationshipManager;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":109,"./util":134,"babel-runtime/core-js/object/create":3,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],121:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.alignment.aspx
//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.horizontalalignmentvalues.aspx
var HORIZONTAL = ['left', 'center', 'right', 'fill', 'justify', 'centerContinuous', 'distributed'];
//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.verticalalignmentvalues.aspx
var VERTICAL = ['top', 'center', 'bottom', 'justify', 'distributed'];

function canon(format) {
	var result = {};

	if (_.has(format, 'horizontal') && _.includes(HORIZONTAL, format.horizontal)) {
		result.horizontal = format.horizontal;
	}
	if (_.has(format, 'vertical') && _.includes(VERTICAL, format.vertical)) {
		result.vertical = format.vertical;
	}
	if (_.has(format, 'indent')) {
		result.indent = format.indent;
	}
	if (_.has(format, 'justifyLastLine')) {
		result.justifyLastLine = format.justifyLastLine ? 1 : 0;
	}
	if (_.has(format, 'readingOrder') && _.includes([0, 1, 2], format.readingOrder)) {
		result.readingOrder = format.readingOrder;
	}
	if (_.has(format, 'relativeIndent')) {
		result.relativeIndent = format.relativeIndent;
	}
	if (_.has(format, 'shrinkToFit')) {
		result.shrinkToFit = format.shrinkToFit ? 1 : 0;
	}
	if (_.has(format, 'textRotation')) {
		result.textRotation = format.textRotation;
	}
	if (_.has(format, 'wrapText')) {
		result.wrapText = format.wrapText ? 1 : 0;
	}

	return _.isEmpty(result) ? null : result;
}

function merge(formatTo, formatFrom) {
	return _.assign(formatTo, formatFrom);
}

function saveFormat(format) {
	return toXMLString({
		name: 'alignment',
		attributes: _.toPairs(format)
	});
}

module.exports = {
	canon: canon,
	merge: merge,
	saveFormat: saveFormat
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109}],122:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var formatUtils = require('./utils');
var toXMLString = require('../XMLString');

var BORDERS = ['left', 'right', 'top', 'bottom', 'diagonal'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.borders.aspx

var Borders = function (_StylePart) {
	(0, _inherits3.default)(Borders, _StylePart);

	function Borders(styles) {
		(0, _classCallCheck3.default)(this, Borders);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Borders.__proto__ || (0, _getPrototypeOf2.default)(Borders)).call(this, styles, 'borders', 'border'));

		_this.init();
		_this.lastId = _this.formats.length;
		return _this;
	}

	(0, _createClass3.default)(Borders, [{
		key: 'init',
		value: function init() {
			this.formats.push({ format: this.canon({}) });
		}
	}, {
		key: 'canon',
		value: function canon(format) {
			return Borders.canon(format);
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			formatTo = formatTo || {};

			if (formatFrom) {
				_.forEach(BORDERS, function (name) {
					var borderFrom = formatFrom[name];

					if (borderFrom && (borderFrom.style || borderFrom.color)) {
						formatTo[name] = {
							style: borderFrom.style,
							color: borderFrom.color
						};
					} else if (!formatTo[name]) {
						formatTo[name] = {};
					}
				});
			}
			return formatTo;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			return Borders.saveFormat(format, styleFormat);
		}
	}], [{
		key: 'canon',
		value: function canon(format) {
			var result = {};

			_.forEach(BORDERS, function (name) {
				var border = format[name];

				if (border) {
					result[name] = {
						style: border.style,
						color: border.color
					};
				} else {
					result[name] = {};
				}
			});
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var children = _.map(BORDERS, function (name) {
				var border = format[name];
				var attributes = void 0;
				var children = void 0;

				if (border) {
					if (border.style) {
						attributes = [['style', border.style]];
					}
					if (border.color) {
						children = [formatUtils.saveColor(border.color)];
					}
				}
				return toXMLString({
					name: name,
					attributes: attributes,
					children: children
				});
			});

			return toXMLString({
				name: 'border',
				children: children
			});
		}
	}]);
	return Borders;
}(StylePart);

module.exports = Borders;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"./stylePart":129,"./utils":132,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],123:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var alignment = require('./alignment');
var protection = require('./protection');
var toXMLString = require('../XMLString');

var ALLOWED_PARTS = ['format', 'fill', 'border', 'font'];
var XLS_NAMES = ['numFmtId', 'fillId', 'borderId', 'fontId'];

var Cells = function (_StylePart) {
	(0, _inherits3.default)(Cells, _StylePart);

	function Cells(styles) {
		(0, _classCallCheck3.default)(this, Cells);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Cells.__proto__ || (0, _getPrototypeOf2.default)(Cells)).call(this, styles, 'cellXfs', 'format'));

		_this.init();
		_this.lastId = _this.formats.length;
		_this.saveEmpty = false;
		return _this;
	}

	(0, _createClass3.default)(Cells, [{
		key: 'init',
		value: function init() {
			this.formats.push({ format: this.canon({}) });
		}
	}, {
		key: 'canon',
		value: function canon(format, flags) {
			var result = {};

			if (format.format) {
				result.format = this.styles.numberFormats.add(format.format);
			}
			if (format.font) {
				result.font = this.styles.fonts.add(format.font);
			}
			if (format.pattern) {
				result.fill = this.styles.fills.add(format.pattern, null, { fillType: 'pattern' });
			} else if (format.gradient) {
				result.fill = this.styles.fills.add(format.gradient, null, { fillType: 'gradient' });
			} else if (flags && flags.merge && format.fill) {
				result.fill = this.styles.fills.add(format.fill, null, flags);
			}
			if (format.border) {
				result.border = this.styles.borders.add(format.border);
			}
			var alignmentValue = alignment.canon(format);
			if (alignmentValue) {
				result.alignment = alignmentValue;
			}
			var protectionValue = protection.canon(format);
			if (protectionValue) {
				result.protection = protectionValue;
			}
			if (format.fillOut) {
				result.fillOut = format.fillOut;
			}
			return result;
		}
	}, {
		key: 'fullGet',
		value: function fullGet(format) {
			var result = {};

			if (this.getId(format)) {
				format = this.get(format);
				if (format.format) {
					result.format = this.styles.numberFormats.get(format.format);
				}
				if (format.font) {
					result.font = _.clone(this.styles.fonts.get(format.font));
				}
				if (format.fill) {
					result.fill = _.clone(this.styles.fills.get(format.fill));
				}
				if (format.border) {
					result.border = _.clone(this.styles.borders.get(format.border));
				}
				if (format.alignment) {
					result.alignment = _.clone(format.alignment);
				}
				if (format.protection) {
					result.protection = _.clone(format.protection);
				}
			} else {
				result = this.canon(format);
			}
			return result;
		}
	}, {
		key: 'cutVisible',
		value: function cutVisible(format) {
			var result = {};

			if (format.format) {
				result.format = format.format;
			}
			if (format.font) {
				result.font = format.font;
			}
			if (format.alignment) {
				result.alignment = format.alignment;
			}
			if (format.protection) {
				result.protection = format.protection;
			}
			return result;
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			if (formatTo.format || formatFrom.format) {
				formatTo.format = this.styles.numberFormats.merge(formatTo.format, formatFrom.format);
			}
			if (formatTo.font || formatFrom.font) {
				formatTo.font = this.styles.fonts.merge(formatTo.font, formatFrom.font);
			}
			if (formatTo.fill || formatFrom.fill) {
				formatTo.fill = this.styles.fills.merge(formatTo.fill, formatFrom.fill);
			}
			if (formatTo.border || formatFrom.border) {
				formatTo.border = this.styles.borders.merge(formatTo.border, formatFrom.border);
			}
			if (formatTo.alignment || formatFrom.alignment) {
				formatTo.alignment = alignment.merge(formatTo.alignment, formatFrom.alignment);
			}
			if (formatTo.protection || formatFrom.protection) {
				formatTo.protection = protection.merge(formatTo.protection, formatFrom.protection);
			}
			return formatTo;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var styles = this.styles;
			var attributes = [];
			var children = [];

			if (format.alignment) {
				children.push(alignment.saveFormat(format.alignment));
				attributes.push(['applyAlignment', 'true']);
			}
			if (format.protection) {
				children.push(protection.saveFormat(format.protection));
				attributes.push(['applyProtection', 'true']);
			}

			_.forEach(format, function (value, key) {
				if (_.includes(ALLOWED_PARTS, key)) {
					var xlsName = XLS_NAMES[_.indexOf(ALLOWED_PARTS, key)];

					if (key === 'format') {
						attributes.push([xlsName, styles.numberFormats.getId(value)]);
						attributes.push(['applyNumberFormat', 'true']);
					} else if (key === 'fill') {
						attributes.push([xlsName, styles.fills.getId(value)]);
						attributes.push(['applyFill', 'true']);
					} else if (key === 'border') {
						attributes.push([xlsName, styles.borders.getId(value)]);
						attributes.push(['applyBorder', 'true']);
					} else if (key === 'font') {
						attributes.push([xlsName, styles.fonts.getId(value)]);
						attributes.push(['applyFont', 'true']);
					}
				}
			});

			return toXMLString({
				name: 'xf',
				attributes: attributes,
				children: children
			});
		}
	}]);
	return Cells;
}(StylePart);

module.exports = Cells;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"./alignment":121,"./protection":128,"./stylePart":129,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],124:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var toXMLString = require('../XMLString');

var PATTERN_TYPES = ['none', 'solid', 'darkGray', 'mediumGray', 'lightGray', 'gray125', 'gray0625', 'darkHorizontal', 'darkVertical', 'darkDown', 'darkUp', 'darkGrid', 'darkTrellis', 'lightHorizontal', 'lightVertical', 'lightDown', 'lightUp', 'lightGrid', 'lightTrellis'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fills.aspx

var Fills = function (_StylePart) {
	(0, _inherits3.default)(Fills, _StylePart);

	function Fills(styles) {
		(0, _classCallCheck3.default)(this, Fills);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Fills.__proto__ || (0, _getPrototypeOf2.default)(Fills)).call(this, styles, 'fills', 'fill'));

		_this.init();
		_this.lastId = _this.formats.length;
		return _this;
	}

	(0, _createClass3.default)(Fills, [{
		key: 'init',
		value: function init() {
			this.formats.push({ format: this.canon({ type: 'none' }, { fillType: 'pattern' }) }, { format: this.canon({ type: 'gray125' }, { fillType: 'pattern' }) });
		}
	}, {
		key: 'canon',
		value: function canon(format, flags) {
			return Fills.canon(format, flags);
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			return formatFrom || formatTo;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			return Fills.saveFormat(format);
		}
	}], [{
		key: 'canon',
		value: function canon(format, flags) {
			var result = {
				fillType: flags.merge ? format.fillType : flags.fillType
			};

			if (result.fillType === 'pattern') {
				var fgColor = (flags.merge ? format.fgColor : format.color) || 'FFFFFFFF';
				var bgColor = (flags.merge ? format.bgColor : format.backColor) || 'FFFFFFFF';
				var patternType = flags.merge ? format.patternType : format.type;

				result.patternType = _.includes(PATTERN_TYPES, patternType) ? patternType : 'solid';
				if (flags.isTable && result.patternType === 'solid') {
					result.fgColor = bgColor;
					result.bgColor = fgColor;
				} else {
					result.fgColor = fgColor;
					result.bgColor = bgColor;
				}
			} else {
				if (_.has(format, 'left')) {
					result.left = format.left || 0;
					result.right = format.right || 0;
					result.top = format.top || 0;
					result.bottom = format.bottom || 0;
				} else {
					result.degree = format.degree || 0;
				}
				result.start = format.start || 'FFFFFFFF';
				result.end = format.end || 'FFFFFFFF';
			}
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var children = void 0;

			if (format.fillType === 'pattern') {
				children = [savePatternFill(format)];
			} else {
				children = [saveGradientFill(format)];
			}

			return toXMLString({
				name: 'fill',
				children: children
			});
		}
	}]);
	return Fills;
}(StylePart);

function savePatternFill(format) {
	var attributes = [['patternType', format.patternType]];
	var children = [toXMLString({
		name: 'fgColor',
		attributes: [['rgb', format.fgColor]]
	}), toXMLString({
		name: 'bgColor',
		attributes: [['rgb', format.bgColor]]
	})];

	return toXMLString({
		name: 'patternFill',
		attributes: attributes,
		children: children
	});
}

function saveGradientFill(format) {
	var attributes = [];
	var children = [];
	var attrs = void 0;

	if (format.degree) {
		attributes.push(['degree', format.degree]);
	} else if (format.left) {
		attributes.push(['type', 'path']);
		attributes.push(['left', format.left]);
		attributes.push(['right', format.right]);
		attributes.push(['top', format.top]);
		attributes.push(['bottom', format.bottom]);
	}

	attrs = [['rgb', format.start]];
	children.push(toXMLString({
		name: 'stop',
		attributes: [['position', 0]],
		children: [toXMLString({ name: 'color', attributes: attrs })]
	}));

	attrs = [['rgb', format.end]];
	children.push(toXMLString({
		name: 'stop',
		attributes: [['position', 1]],
		children: [toXMLString({ name: 'color', attributes: attrs })]
	}));

	return toXMLString({
		name: 'gradientFill',
		attributes: attributes,
		children: children
	});
}

module.exports = Fills;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"./stylePart":129,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],125:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var formatUtils = require('./utils');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.fonts.aspx

var Fonts = function (_StylePart) {
	(0, _inherits3.default)(Fonts, _StylePart);

	function Fonts(styles) {
		(0, _classCallCheck3.default)(this, Fonts);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Fonts.__proto__ || (0, _getPrototypeOf2.default)(Fonts)).call(this, styles, 'fonts', 'font'));

		_this.init();
		_this.lastId = _this.formats.length;
		return _this;
	}

	(0, _createClass3.default)(Fonts, [{
		key: 'init',
		value: function init() {
			this.formats.push({ format: this.canon({}) });
		}
	}, {
		key: 'canon',
		value: function canon(format) {
			return Fonts.canon(format);
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			var result = _.assign(formatTo, formatFrom);

			result.color = formatFrom && formatFrom.color || formatTo && formatTo.color;
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			return Fonts.saveFormat(format, styleFormat);
		}
	}], [{
		key: 'canon',
		value: function canon(format) {
			var result = {};

			if (_.has(format, 'bold')) {
				result.bold = !!format.bold;
			}
			if (_.has(format, 'italic')) {
				result.italic = !!format.italic;
			}
			if (format.superscript) {
				result.vertAlign = 'superscript';
			}
			if (format.subscript) {
				result.vertAlign = 'subscript';
			}
			if (format.underline) {
				if (_.indexOf(['double', 'singleAccounting', 'doubleAccounting'], format.underline) !== -1) {
					result.underline = format.underline;
				} else {
					result.underline = true;
				}
			}
			if (_.has(format, 'strike')) {
				result.strike = !!format.strike;
			}
			if (format.outline) {
				result.outline = true;
			}
			if (format.shadow) {
				result.shadow = true;
			}
			if (format.size) {
				result.size = format.size;
			}
			if (format.color) {
				result.color = format.color;
			}
			if (format.fontName) {
				result.fontName = format.fontName;
			}
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var children = [];
			var attrs = void 0;

			if (format.size) {
				children.push(toXMLString({
					name: 'sz',
					attributes: [['val', format.size]]
				}));
			}
			if (format.fontName) {
				children.push(toXMLString({
					name: 'name',
					attributes: [['val', format.fontName]]
				}));
			}
			if (_.has(format, 'bold')) {
				if (format.bold) {
					children.push(toXMLString({ name: 'b' }));
				} else {
					children.push(toXMLString({ name: 'b', attributes: [['val', 0]] }));
				}
			}
			if (_.has(format, 'italic')) {
				if (format.italic) {
					children.push(toXMLString({ name: 'i' }));
				} else {
					children.push(toXMLString({ name: 'i', attributes: [['val', 0]] }));
				}
			}
			if (format.vertAlign) {
				children.push(toXMLString({
					name: 'vertAlign',
					attributes: [['val', format.vertAlign]]
				}));
			}
			if (format.underline) {
				attrs = null;

				if (format.underline !== true) {
					attrs = [['val', format.underline]];
				}
				children.push(toXMLString({
					name: 'u',
					attributes: attrs
				}));
			}
			if (_.has(format, 'strike')) {
				if (format.strike) {
					children.push(toXMLString({ name: 'strike' }));
				} else {
					children.push(toXMLString({ name: 'strike', attributes: [['val', 0]] }));
				}
			}
			if (format.shadow) {
				children.push(toXMLString({ name: 'shadow' }));
			}
			if (format.outline) {
				children.push(toXMLString({ name: 'outline' }));
			}
			if (format.color) {
				children.push(formatUtils.saveColor(format.color));
			}

			return toXMLString({
				name: 'font',
				children: children
			});
		}
	}]);
	return Fonts;
}(StylePart);

module.exports = Fonts;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"./stylePart":129,"./utils":132,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],126:[function(require,module,exports){
(function (global){
'use strict';

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var NumberFormats = require('./numberFormats');
var Fonts = require('./fonts');
var Fills = require('./fills');
var Borders = require('./borders');
var Cells = require('./cells');
var Tables = require('./tables');
var TableElements = require('./tableElements');
var toXMLString = require('../XMLString');

var Styles = function () {
	function Styles(common) {
		(0, _classCallCheck3.default)(this, Styles);

		this.objectId = common.uniqueId('Styles');
		this.numberFormats = new NumberFormats(this);
		this.fonts = new Fonts(this);
		this.fills = new Fills(this);
		this.borders = new Borders(this);
		this.cells = new Cells(this);
		this.tableElements = new TableElements(this);
		this.tables = new Tables(this);
		this.defaultTableStyle = '';
	}

	(0, _createClass3.default)(Styles, [{
		key: 'addFormat',
		value: function addFormat(format, name) {
			return this.cells.add(format, name);
		}
	}, {
		key: '_get',
		value: function _get(name) {
			return this.cells.get(name);
		}
	}, {
		key: '_getId',
		value: function _getId(name) {
			return this.cells.getId(name);
		}
	}, {
		key: '_addInvisibleFormat',
		value: function _addInvisibleFormat(format) {
			var style = this.cells.cutVisible(this.cells.fullGet(format));

			if (!_.isEmpty(style)) {
				return this.addFormat(style);
			}
		}
	}, {
		key: '_merge',
		value: function _merge(columnFormat, rowFormat, cellFormat) {
			var count = Number(Boolean(columnFormat)) + Number(Boolean(rowFormat)) + Number(Boolean(cellFormat));

			if (count === 0) {
				return null;
			} else if (count === 1) {
				return this.cells.add(columnFormat || rowFormat || cellFormat);
			} else {
				var format = {};

				if (columnFormat) {
					format = this.cells.merge(format, this.cells.fullGet(columnFormat));
				}
				if (rowFormat) {
					format = this.cells.merge(format, this.cells.fullGet(rowFormat));
				}
				if (cellFormat) {
					format = this.cells.merge(format, this.cells.fullGet(cellFormat));
				}
				return this.cells.add(format, null, { merge: true });
			}
		}
	}, {
		key: 'addFontFormat',
		value: function addFontFormat(format, name) {
			return this.fonts.add(format, name);
		}
	}, {
		key: 'addBorderFormat',
		value: function addBorderFormat(format, name) {
			return this.borders.add(format, name);
		}
	}, {
		key: 'addPatternFormat',
		value: function addPatternFormat(format, name) {
			return this.fills.add(format, name, { fillType: 'pattern' });
		}
	}, {
		key: 'addGradientFormat',
		value: function addGradientFormat(format, name) {
			return this.fills.add(format, name, { fillType: 'gradient' });
		}
	}, {
		key: 'addNumberFormat',
		value: function addNumberFormat(format, name) {
			return this.numberFormats.add(format, name);
		}
	}, {
		key: 'addTableFormat',
		value: function addTableFormat(format, name) {
			return this.tables.add(format, name);
		}
	}, {
		key: 'addTableElementFormat',
		value: function addTableElementFormat(format, name) {
			return this.tableElements.add(format, name);
		}
	}, {
		key: 'setDefaultTableStyle',
		value: function setDefaultTableStyle(name) {
			this.tables.defaultTableStyle = name;
		}
	}, {
		key: 'save',
		value: function save() {
			return toXMLString({
				name: 'styleSheet',
				ns: 'spreadsheetml',
				children: [this.numberFormats.save(), this.fonts.save(), this.fills.save(), this.borders.save(), this.cells.save(), this.tableElements.save(), this.tables.save()]
			});
		}
	}]);
	return Styles;
}();

module.exports = Styles;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"./borders":122,"./cells":123,"./fills":124,"./fonts":125,"./numberFormats":127,"./tableElements":130,"./tables":131,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],127:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var toXMLString = require('../XMLString');

var PREDEFINED = {
	date: 14, //mm-dd-yy
	time: 21 //h:mm:ss
};

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.numberingformats.aspx

var NumberFormats = function (_StylePart) {
	(0, _inherits3.default)(NumberFormats, _StylePart);

	function NumberFormats(styles) {
		(0, _classCallCheck3.default)(this, NumberFormats);

		var _this = (0, _possibleConstructorReturn3.default)(this, (NumberFormats.__proto__ || (0, _getPrototypeOf2.default)(NumberFormats)).call(this, styles, 'numFmts', 'numberFormat'));

		_this.init();
		_this.lastId = 164;
		return _this;
	}

	(0, _createClass3.default)(NumberFormats, [{
		key: 'init',
		value: function init() {
			var _this2 = this;

			_.forEach(PREDEFINED, function (formatId, format) {
				_this2.formatsByNames[format] = {
					formatId: formatId,
					format: format
				};
			});
		}
	}, {
		key: 'canon',
		value: function canon(format) {
			return NumberFormats.canon(format);
		}
	}, {
		key: 'merge',
		value: function merge(formatTo, formatFrom) {
			return formatFrom || formatTo;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			return NumberFormats.saveFormat(format, styleFormat);
		}
	}], [{
		key: 'canon',
		value: function canon(format) {
			return format;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			var attributes = [['numFmtId', styleFormat.formatId], ['formatCode', format]];

			return toXMLString({
				name: 'numFmt',
				attributes: attributes
			});
		}
	}]);
	return NumberFormats;
}(StylePart);

module.exports = NumberFormats;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"./stylePart":129,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],128:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.protection.aspx

function canon(format) {
	var result = {};

	if (_.has(format, 'locked')) {
		result.locked = format.locked ? 1 : 0;
	}
	if (_.has(format, 'hidden')) {
		result.hidden = format.hidden ? 1 : 0;
	}

	return _.isEmpty(result) ? null : result;
}

function merge(formatTo, formatFrom) {
	return _.assign(formatTo, formatFrom);
}

function saveFormat(format) {
	return toXMLString({
		name: 'protection',
		attributes: _.toPairs(format)
	});
}

module.exports = {
	canon: canon,
	merge: merge,
	saveFormat: saveFormat
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109}],129:[function(require,module,exports){
(function (global){
'use strict';

var _stringify = require('babel-runtime/core-js/json/stringify');

var _stringify2 = _interopRequireDefault(_stringify);

var _create = require('babel-runtime/core-js/object/create');

var _create2 = _interopRequireDefault(_create);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

var StylePart = function () {
	function StylePart(styles, saveName, formatName) {
		(0, _classCallCheck3.default)(this, StylePart);

		this.styles = styles;
		this.saveName = saveName;
		this.formatName = formatName;
		this.lastName = 1;
		this.lastId = 0;
		this.saveEmpty = true;
		this.formats = [];
		this.formatsByData = (0, _create2.default)(null);
		this.formatsByNames = (0, _create2.default)(null);
	}

	(0, _createClass3.default)(StylePart, [{
		key: 'add',
		value: function add(format, name, flags) {
			if (name && this.formatsByNames[name]) {
				var _canonFormat = this.canon(format, flags);
				var _stringFormat = _.isObject(_canonFormat) ? (0, _stringify2.default)(_canonFormat) : _canonFormat;

				if (_stringFormat !== this.formatsByNames[name].stringFormat) {
					this._add(_canonFormat, _stringFormat, name);
				}
				return name;
			}

			//first argument is format name
			if (!name && _.isString(format) && this.formatsByNames[format]) {
				return format;
			}

			var canonFormat = this.canon(format, flags);
			var stringFormat = _.isObject(canonFormat) ? (0, _stringify2.default)(canonFormat) : canonFormat;
			var styleFormat = this.formatsByData[stringFormat];

			if (!styleFormat) {
				styleFormat = this._add(canonFormat, stringFormat, name);
			} else if (name && !this.formatsByNames[name]) {
				styleFormat.name = name;
				this.formatsByNames[name] = styleFormat;
			}
			return styleFormat.name;
		}
	}, {
		key: '_add',
		value: function _add(canonFormat, stringFormat, name) {
			name = name || this.formatName + this.lastName++;

			var styleFormat = {
				name: name,
				formatId: this.lastId++,
				format: canonFormat,
				stringFormat: stringFormat
			};

			this.formats.push(styleFormat);
			this.formatsByData[stringFormat] = styleFormat;
			this.formatsByNames[name] = styleFormat;

			return styleFormat;
		}
	}, {
		key: 'canon',
		value: function canon(format) {
			return format;
		}
	}, {
		key: 'get',
		value: function get(format) {
			if (_.isString(format)) {
				var styleFormat = this.formatsByNames[format];

				return styleFormat ? styleFormat.format : format;
			}
			return format;
		}
	}, {
		key: 'getId',
		value: function getId(name) {
			var styleFormat = this.formatsByNames[name];

			return styleFormat ? styleFormat.formatId : null;
		}
	}, {
		key: 'save',
		value: function save() {
			var _this = this;

			if (this.saveEmpty !== false || this.formats.length) {
				var attributes = [['count', this.formats.length]];
				var children = _.map(this.formats, function (format) {
					return _this.saveFormat(format.format, format);
				});

				this.saveCollectionExt(attributes, children);

				return toXMLString({
					name: this.saveName,
					attributes: attributes,
					children: children
				});
			}
			return '';
		}
	}, {
		key: 'saveCollectionExt',
		value: function saveCollectionExt() {}
	}, {
		key: 'saveFormat',
		value: function saveFormat() {
			return '';
		}
	}]);
	return StylePart;
}();

module.exports = StylePart;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"babel-runtime/core-js/json/stringify":1,"babel-runtime/core-js/object/create":3,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],130:[function(require,module,exports){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var StylePart = require('./stylePart');
var NumberFormats = require('./numberFormats');
var Fonts = require('./fonts');
var Fills = require('./fills');
var Borders = require('./borders');
var alignment = require('./alignment');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.differentialformats.aspx

var TableElements = function (_StylePart) {
	(0, _inherits3.default)(TableElements, _StylePart);

	function TableElements(styles) {
		(0, _classCallCheck3.default)(this, TableElements);
		return (0, _possibleConstructorReturn3.default)(this, (TableElements.__proto__ || (0, _getPrototypeOf2.default)(TableElements)).call(this, styles, 'dxfs', 'tableElement'));
	}

	(0, _createClass3.default)(TableElements, [{
		key: 'canon',
		value: function canon(format) {
			var result = {};

			if (format.format) {
				result.format = NumberFormats.canon(format.format);
			}
			if (format.font) {
				result.font = Fonts.canon(format.font);
			}
			if (format.pattern) {
				result.fill = Fills.canon(format.pattern, { fillType: 'pattern', isTable: true });
			} else if (format.gradient) {
				result.fill = Fills.canon(format.gradient, { fillType: 'gradient' });
			}
			if (format.border) {
				result.border = Borders.canon(format.border);
			}
			result.alignment = alignment.canon(format);
			return result;
		}
	}, {
		key: 'saveFormat',
		value: function saveFormat(format) {
			var children = [];

			if (format.font) {
				children.push(Fonts.saveFormat(format.font));
			}
			if (format.fill) {
				children.push(Fills.saveFormat(format.fill));
			}
			if (format.border) {
				children.push(Borders.saveFormat(format.border));
			}
			if (format.format) {
				children.push(NumberFormats.saveFormat(format.format));
			}
			if (format.alignment && format.alignment.length) {
				children.push(alignment.save(format.alignment));
			}

			return toXMLString({
				name: 'dxf',
				children: children
			});
		}
	}]);
	return TableElements;
}(StylePart);

module.exports = TableElements;

},{"../XMLString":109,"./alignment":121,"./borders":122,"./fills":124,"./fonts":125,"./numberFormats":127,"./stylePart":129,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],131:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var StylePart = require('./stylePart');
var toXMLString = require('../XMLString');

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestylevalues.aspx
var ELEMENTS = ['wholeTable', 'headerRow', 'totalRow', 'firstColumn', 'lastColumn', 'firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe', 'firstHeaderCell', 'lastHeaderCell', 'firstTotalCell', 'lastTotalCell'];
var SIZED_ELEMENTS = ['firstRowStripe', 'secondRowStripe', 'firstColumnStripe', 'secondColumnStripe'];

//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyles.aspx

var Tables = function (_StylePart) {
	(0, _inherits3.default)(Tables, _StylePart);

	function Tables(styles) {
		(0, _classCallCheck3.default)(this, Tables);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Tables.__proto__ || (0, _getPrototypeOf2.default)(Tables)).call(this, styles, 'tableStyles', 'table'));

		_this.saveEmpty = false;
		return _this;
	}

	(0, _createClass3.default)(Tables, [{
		key: 'canon',
		value: function canon(format) {
			var result = {};
			var styles = this.styles;

			_.forEach(format, function (value, key) {
				if (_.includes(ELEMENTS, key)) {
					var style = void 0;
					var size = null;

					if (value.style) {
						style = styles.tableElements.add(value.style);
						if (value.size > 1 && _.includes(SIZED_ELEMENTS, key)) {
							size = value.size;
						}
					} else {
						style = styles.tableElements.add(value);
					}
					result[key] = {
						style: style,
						size: size
					};
				}
			});

			return result;
		}
	}, {
		key: 'saveCollectionExt',
		value: function saveCollectionExt(attributes) {
			if (this.styles.defaultTableStyle) {
				attributes.push(['defaultTableStyle', this.styles.defaultTableStyle]);
			}
		}
		//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.tablestyleelement.aspx

	}, {
		key: 'saveFormat',
		value: function saveFormat(format, styleFormat) {
			var styles = this.styles;
			var attributes = [['name', styleFormat.name], ['pivot', 0]];
			var children = [];

			_.forEach(format, function (value, key) {
				var style = value.style;
				var size = value.size;
				var attributes = [['type', key], ['dxfId', styles.tableElements.getId(style)]];

				if (size) {
					attributes.push(['size', size]);
				}

				children.push(toXMLString({
					name: 'tableStyleElement',
					attributes: attributes
				}));
			});
			attributes.push(['count', children.length]);

			return toXMLString({
				name: 'tableStyle',
				attributes: attributes,
				children: children
			});
		}
	}]);
	return Tables;
}(StylePart);

module.exports = Tables;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"./stylePart":129,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],132:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

function saveColor(color) {
	if (_.isString(color)) {
		return toXMLString({
			name: 'color',
			attributes: [['rgb', color]]
		});
	} else {
		var attributes = [];
		if (!_.isUndefined(color.tint)) {
			attributes.push(['tint', color.tint]);
		}
		if (!_.isUndefined(color.auto)) {
			attributes.push(['auto', !!color.auto]);
		}
		if (!_.isUndefined(color.theme)) {
			attributes.push(['theme', color.theme]);
		}

		return toXMLString({
			name: 'color',
			attributes: attributes
		});
	}
}

module.exports = {
	saveColor: saveColor
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109}],133:[function(require,module,exports){
(function (global){
'use strict';

var _typeof2 = require('babel-runtime/helpers/typeof');

var _typeof3 = _interopRequireDefault(_typeof2);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');
var toXMLString = require('./XMLString');

var Table = function () {
	function Table(worksheet, config) {
		(0, _classCallCheck3.default)(this, Table);

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

	(0, _createClass3.default)(Table, [{
		key: 'end',
		value: function end() {
			return this.worksheet;
		}
	}, {
		key: 'setReferenceRange',
		value: function setReferenceRange(beginCell, endCell) {
			this.beginCell = util.canonCell(beginCell);
			this.endCell = util.canonCell(endCell);
			return this;
		}
	}, {
		key: 'addTotalRow',
		value: function addTotalRow(totalRow) {
			this.totalRow = totalRow;
			this.totalsRowCount = 1;
			return this;
		}
	}, {
		key: 'setTheme',
		value: function setTheme(theme) {
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

	}, {
		key: 'setSortState',
		value: function setSortState(state) {
			this.sortState = state;
			return this;
		}
	}, {
		key: '_prepare',
		value: function _prepare(worksheetData) {
			var _this = this;

			var SUB_TOTAL_FUNCTIONS = ['average', 'countNums', 'count', 'max', 'min', 'stdDev', 'sum', 'var'];
			var SUB_TOTAL_NUMS = [101, 102, 103, 104, 105, 107, 109, 110];

			if (this.totalRow) {
				(function () {
					var tableName = _this.name;
					var beginCell = util.letterToPosition(_this.beginCell);
					var endCell = util.letterToPosition(_this.endCell);
					var firstRow = beginCell.y - 1;
					var firstColumn = beginCell.x - 1;
					var lastRow = endCell.y - 1;
					var headerRow = worksheetData[firstRow] || [];
					var totalRow = worksheetData[lastRow + 1];

					if (!totalRow) {
						totalRow = [];
						worksheetData[lastRow + 1] = totalRow;
					}

					_.forEach(_this.totalRow, function (cell, i) {
						var headerValue = headerRow[firstColumn + i];
						var funcIndex = void 0;

						if ((typeof headerValue === 'undefined' ? 'undefined' : (0, _typeof3.default)(headerValue)) === 'object') {
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
									value: 'SUBTOTAL(' + SUB_TOTAL_NUMS[funcIndex] + ',' + tableName + '[' + headerValue + '])',
									type: 'formula'
								};
							}
						}
					});
				})();
			}
		}
		//https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.table.aspx

	}, {
		key: '_save',
		value: function _save() {
			var attributes = [['id', this.tableId], ['name', this.name], ['displayName', this.displayName]];
			var children = [];
			var end = util.letterToPosition(this.endCell);
			var ref = this.beginCell + ':' + util.positionToLetter(end.x, end.y + this.totalsRowCount);

			attributes.push(['ref', ref]);
			attributes.push(['totalsRowCount', this.totalsRowCount]);
			attributes.push(['headerRowCount', this.headerRowCount]);

			children.push(saveAutoFilter(this.beginCell, this.endCell));
			children.push(saveTableColumns(this.totalRow));
			children.push(saveTableStyleInfo(this.common, this.themeStyle));

			return toXMLString({
				name: 'table',
				ns: 'spreadsheetml',
				attributes: attributes,
				children: children
			});
		}
	}]);
	return Table;
}();

function saveAutoFilter(beginCell, endCell) {
	return toXMLString({
		name: 'autoFilter',
		attributes: ['ref', beginCell + ':' + endCell]
	});
}

function saveTableColumns(totalRow) {
	var attributes = [['count', totalRow.length]];
	var children = _.map(totalRow, function (cell, index) {
		var attributes = [['id', index + 1], ['name', cell.name]];

		if (cell.totalsRowFunction) {
			attributes.push(['totalsRowFunction', cell.totalsRowFunction]);
		}
		if (cell.totalsRowLabel) {
			attributes.push(['totalsRowLabel', cell.totalsRowLabel]);
		}

		return toXMLString({
			name: 'tableColumn',
			attributes: attributes
		});
	});

	return toXMLString({
		name: 'tableColumns',
		attributes: attributes,
		children: children
	});
}

function saveTableStyleInfo(common, themeStyle) {
	var attributes = [['name', themeStyle]];
	var format = common.styles.tables.get(themeStyle);
	var isRowStripes = false;
	var isColumnStripes = false;
	var isFirstColumn = false;
	var isLastColumn = false;

	if (format) {
		isRowStripes = format.firstRowStripe || format.secondRowStripe;
		isColumnStripes = format.firstColumnStripe || format.secondColumnStripe;
		isFirstColumn = format.firstColumn;
		isLastColumn = format.lastColumn;
	}
	attributes.push(['showRowStripes', isRowStripes ? '1' : '0'], ['showColumnStripes', isColumnStripes ? '1' : '0'], ['showFirstColumn', isFirstColumn ? '1' : '0'], ['showLastColumn', isLastColumn ? '1' : '0']);

	return toXMLString({
		name: 'tableStyleInfo',
		attributes: attributes
	});
}

module.exports = Table;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":109,"./util":134,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/typeof":13}],134:[function(require,module,exports){
(function (global){
'use strict';

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var alphabet = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';
var LETTER_REFS = {};

function positionToLetter(x, y) {
	var result = LETTER_REFS[x];

	if (!result) {
		var string = '';
		var num = x;
		var index = void 0;

		do {
			index = (num - 1) % 26;
			string = alphabet[index] + string;
			num = (num - (index + 1)) / 26;
		} while (num > 0);

		LETTER_REFS[x] = string;
		result = string;
	}
	return result + (y || '');
}

function letterToPosition(cell) {
	var x = 0;
	var y = 0;
	var i = void 0;
	var len = void 0;
	var charCode = void 0;

	for (i = 0, len = cell.length; i < len; i++) {
		charCode = cell.charCodeAt(i);
		if (charCode >= 65) {
			x = x * 26 + charCode - 64;
		} else {
			y = parseInt(cell.slice(i), 10);
			break;
		}
	}
	return {
		x: x || 1,
		y: y || 1
	};
}

module.exports = {
	pixelsToEMUs: function pixelsToEMUs(pixels) {
		return Math.round(pixels * 914400 / 96);
	},

	canonCell: function canonCell(cell) {
		if (_.isObject(cell)) {
			return positionToLetter(cell.c || 1, cell.r || 1);
		}
		return cell;
	},

	positionToLetter: positionToLetter,
	letterToPosition: letterToPosition,

	xmlPrefix: '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n',

	schemas: {
		'worksheet': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet',
		'sharedStrings': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings',
		'stylesheet': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles',
		'relationships': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
		'relationshipPackage': 'http://schemas.openxmlformats.org/package/2006/relationships',
		'contentTypes': 'http://schemas.openxmlformats.org/package/2006/content-types',
		'spreadsheetml': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
		'markupCompat': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
		'x14ac': 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac',
		'officeDocument': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument',
		'package': 'http://schemas.openxmlformats.org/package/2006/relationships',
		'table': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/table',
		'spreadsheetDrawing': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
		'drawing': 'http://schemas.openxmlformats.org/drawingml/2006/main',
		'drawingRelationship': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing',
		'image': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
		'chart': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart',
		'hyperlink': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink'
	}
};

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{}],135:[function(require,module,exports){
(function (global){
'use strict';

var _typeof2 = require('babel-runtime/helpers/typeof');

var _typeof3 = _interopRequireDefault(_typeof2);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('./util');
var Common = require('./common');
var Relations = require('./relations');
var Worksheet = require('./worksheet');
var toXMLString = require('./XMLString');

var Workbook = function () {
	function Workbook() {
		(0, _classCallCheck3.default)(this, Workbook);

		this.common = new Common();
		this.styles = this.common.styles;
		this.relations = new Relations(this.common);

		this.objectId = this.common.uniqueId('Workbook');
	}

	(0, _createClass3.default)(Workbook, [{
		key: 'addWorksheet',
		value: function addWorksheet(config) {
			config = _.defaults(config, {
				name: this.common.getNewWorksheetDefaultName()
			});
			var worksheet = new Worksheet(this, config);
			this.common.addWorksheet(worksheet);
			this.relations.addRelation(worksheet, 'worksheet');

			return worksheet;
		}
	}, {
		key: 'addFormat',
		value: function addFormat(format, name) {
			return this.styles.addFormat(format, name);
		}
	}, {
		key: 'addFontFormat',
		value: function addFontFormat(format, name) {
			return this.styles.addFontFormat(format, name);
		}
	}, {
		key: 'addBorderFormat',
		value: function addBorderFormat(format, name) {
			return this.styles.addBorderFormat(format, name);
		}
	}, {
		key: 'addPatternFormat',
		value: function addPatternFormat(format, name) {
			return this.styles.addPatternFormat(format, name);
		}
	}, {
		key: 'addGradientFormat',
		value: function addGradientFormat(format, name) {
			return this.styles.addGradientFormat(format, name);
		}
	}, {
		key: 'addNumberFormat',
		value: function addNumberFormat(format, name) {
			return this.styles.addNumberFormat(format, name);
		}
	}, {
		key: 'addTableFormat',
		value: function addTableFormat(format, name) {
			return this.styles.addTableFormat(format, name);
		}
	}, {
		key: 'addTableElementFormat',
		value: function addTableElementFormat(format, name) {
			return this.styles.addTableElementFormat(format, name);
		}
	}, {
		key: 'setDefaultTableStyle',
		value: function setDefaultTableStyle(name) {
			this.styles.setDefaultTableStyle(name);
			return this;
		}
	}, {
		key: 'addImage',
		value: function addImage(data, type, name) {
			return this.common.addImage(data, type, name);
		}
	}, {
		key: '_generateFiles',
		value: function _generateFiles(zip, canStream) {
			prepareWorksheets(this.common);

			saveWorksheets(zip, canStream, this.common);
			saveTables(zip, this.common);
			saveImages(zip, this.common);
			saveDrawings(zip, this.common);
			saveStyles(zip, this.relations, this.styles);
			saveSharedStrings(zip, canStream, this.relations, this.common);
			zip.file('[Content_Types].xml', createContentTypes(this.common));
			zip.file('_rels/.rels', createWorkbookRelationship());
			zip.file('xl/workbook.xml', this._save());
			zip.file('xl/_rels/workbook.xml.rels', this.relations.save());
		}
	}, {
		key: '_save',
		value: function _save() {
			return toXMLString({
				name: 'workbook',
				ns: 'spreadsheetml',
				attributes: [['xmlns:r', util.schemas.relationships]],
				children: [bookViewsXML(this.common), sheetsXML(this.relations, this.common), definedNamesXML(this.common)]
			});
		}
	}]);
	return Workbook;
}();

function bookViewsXML(common) {
	var activeTab = 0;
	var activeWorksheetId = void 0;

	if (common.activeWorksheet) {
		activeWorksheetId = common.activeWorksheet.objectId;

		activeTab = Math.max(activeTab, _.findIndex(common.worksheets, function (worksheet) {
			return worksheet.objectId === activeWorksheetId;
		}));
	}

	return toXMLString({
		name: 'bookViews',
		children: [toXMLString({
			name: 'workbookView',
			attributes: [['activeTab', activeTab]]
		})]
	});
}

function sheetsXML(relations, common) {
	var maxWorksheetNameLength = 31;
	var children = _.map(common.worksheets, function (worksheet, index) {
		// Microsoft Excel (2007, 2013) do not allow worksheet names longer than 31 characters
		// if the worksheet name is longer, Excel displays an 'Excel found unreadable content...' popup when opening the file
		if (worksheet.name.length > maxWorksheetNameLength) {
			throw 'Microsoft Excel requires work sheet names to be less than ' + (maxWorksheetNameLength + 1) + ' characters long, work sheet name "' + worksheet.name + '" is ' + worksheet.name.length + ' characters long';
		}

		return toXMLString({
			name: 'sheet',
			attributes: [['name', worksheet.name], ['sheetId', index + 1], ['r:id', relations.getRelationshipId(worksheet)], ['state', worksheet.getState()]]
		});
	});

	return toXMLString({
		name: 'sheets',
		children: children
	});
}

function definedNamesXML(common) {
	var isPrintTitles = _.some(common.worksheets, function (worksheet) {
		return worksheet._printTitles && (worksheet._printTitles.topTo >= 0 || worksheet._printTitles.leftTo >= 0);
	});

	if (isPrintTitles) {
		var _ret = function () {
			var children = [];

			_.forEach(common.worksheets, function (worksheet, index) {
				var entry = worksheet._printTitles;

				if (entry && (entry.topTo >= 0 || entry.leftTo >= 0)) {
					var name = worksheet.name;
					var value = '';

					if (entry.topTo >= 0) {
						value = name + '!$' + (entry.topFrom + 1) + ':$' + (entry.topTo + 1);
						if (entry.leftTo >= 0) {
							value += ',';
						}
					}
					if (entry.leftTo >= 0) {
						value += name + '!$' + util.positionToLetter(entry.leftFrom + 1) + ':$' + util.positionToLetter(entry.leftTo + 1);
					}

					children.push(toXMLString({
						name: 'definedName',
						value: value,
						attributes: [['name', '_xlnm.Print_Titles'], ['localSheetId', index]]
					}));
				}
			});

			return {
				v: toXMLString({
					name: 'definedNames',
					children: children
				})
			};
		}();

		if ((typeof _ret === 'undefined' ? 'undefined' : (0, _typeof3.default)(_ret)) === "object") return _ret.v;
	}
	return '';
}

function prepareWorksheets(common) {
	_.forEach(common.worksheets, function (worksheet) {
		worksheet._prepare();
	});
}

function saveWorksheets(zip, canStream, common) {
	_.forEach(common.worksheets, function (worksheet) {
		zip.file(worksheet.path, worksheet._save(canStream));
		zip.file(worksheet.relationsPath, worksheet.relations.save());
	});
}

function saveTables(zip, common) {
	_.forEach(common.tables, function (table) {
		zip.file(table.path, table._save());
	});
}

function saveImages(zip, common) {
	_.forEach(common.getImages(), function (image) {
		zip.file(image.path, image.data, { base64: true, binary: true });
		image.data = null;
	});
	common.removeImages();
}

function saveDrawings(zip, common) {
	_.forEach(common.drawings, function (drawing) {
		zip.file(drawing.path, drawing.save());
		zip.file(drawing.relationsPath, drawing.relations.save());
	});
}

function saveStyles(zip, relations, styles) {
	relations.addRelation(styles, 'stylesheet');
	zip.file('xl/styles.xml', styles.save());
}

function saveSharedStrings(zip, canStream, relations, common) {
	if (!common.sharedStrings.isEmpty()) {
		relations.addRelation(common.sharedStrings, 'sharedStrings');
		zip.file('xl/sharedStrings.xml', common.sharedStrings.save(canStream));
	}
}

function createContentTypes(common) {
	var children = [];

	children.push(toXMLString({
		name: 'Default',
		attributes: [['Extension', 'rels'], ['ContentType', 'application/vnd.openxmlformats-package.relationships+xml']]
	}));
	children.push(toXMLString({
		name: 'Default',
		attributes: [['Extension', 'xml'], ['ContentType', 'application/xml']]
	}));
	children.push(toXMLString({
		name: 'Override',
		attributes: [['PartName', '/xl/workbook.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml']]
	}));
	if (!common.sharedStrings.isEmpty()) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/sharedStrings.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml']]
		}));
	}
	children.push(toXMLString({
		name: 'Override',
		attributes: [['PartName', '/xl/styles.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml']]
	}));

	_.forEach(common.worksheets, function (worksheet, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/worksheets/sheet' + (index + 1) + '.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml']]
		}));
	});
	_.forEach(common.tables, function (table, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/tables/table' + (index + 1) + '.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml']]
		}));
	});
	_.forEach(common.getExtensions(), function (contentType, extension) {
		children.push(toXMLString({
			name: 'Default',
			attributes: [['Extension', extension], ['ContentType', contentType]]
		}));
	});
	_.forEach(common.drawings, function (drawing, index) {
		children.push(toXMLString({
			name: 'Override',
			attributes: [['PartName', '/xl/drawings/drawing' + (index + 1) + '.xml'], ['ContentType', 'application/vnd.openxmlformats-officedocument.drawing+xml']]
		}));
	});

	return toXMLString({
		name: 'Types',
		ns: 'contentTypes',
		children: children
	});
}

function createWorkbookRelationship() {
	return toXMLString({
		name: 'Relationships',
		ns: 'relationshipPackage',
		children: [toXMLString({
			name: 'Relationship',
			attributes: [['Id', 'rId1'], ['Type', util.schemas.officeDocument], ['Target', 'xl/workbook.xml']]
		})]
	});
}

module.exports = Workbook;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./XMLString":109,"./common":111,"./relations":120,"./util":134,"./worksheet":138,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/typeof":13}],136:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var Tables = require('./tables');
var Drawings = require('../drawings');
var toXMLString = require('../XMLString');

var DrawingsExt = function (_Tables) {
	(0, _inherits3.default)(DrawingsExt, _Tables);

	function DrawingsExt() {
		(0, _classCallCheck3.default)(this, DrawingsExt);

		var _this = (0, _possibleConstructorReturn3.default)(this, (DrawingsExt.__proto__ || (0, _getPrototypeOf2.default)(DrawingsExt)).call(this));

		_this._drawings = null;
		return _this;
	}

	(0, _createClass3.default)(DrawingsExt, [{
		key: 'setImage',
		value: function setImage(image, config) {
			return this._setDrawing(image, config, 'anchor');
		}
	}, {
		key: 'setImageOneCell',
		value: function setImageOneCell(image, config) {
			return this._setDrawing(image, config, 'oneCell');
		}
	}, {
		key: 'setImageAbsolute',
		value: function setImageAbsolute(image, config) {
			return this._setDrawing(image, config, 'absolute');
		}
	}, {
		key: '_setDrawing',
		value: function _setDrawing(image, config, anchorType) {
			var name = void 0;

			if (!this._drawings) {
				this._drawings = new Drawings(this.common);

				this.common.addDrawings(this._drawings);
				this.relations.addRelation(this._drawings, 'drawingRelationship');
			}

			if (_.isObject(image)) {
				name = this.common.addImage(image.data, image.type);
			} else {
				name = image;
			}

			this._drawings.addImage(name, config, anchorType);
			return this;
		}
	}, {
		key: '_insertDrawing',
		value: function _insertDrawing(colIndex, rowIndex, image) {
			var config = void 0;

			if (typeof image === 'string' || image.data) {
				this._setDrawing(image, { c: colIndex + 1, r: rowIndex + 1 }, 'anchor');
			} else {
				config = image.config || {};
				config.cell = { c: colIndex + 1, r: rowIndex + 1 };

				this._setDrawing(image.image, config, 'anchor');
			}
		}
	}, {
		key: '_saveDrawing',
		value: function _saveDrawing() {
			if (this._drawings) {
				return toXMLString({
					name: 'drawing',
					attributes: [['r:id', this.relations.getRelationshipId(this._drawings)]]
				});
			}
			return '';
		}
	}]);
	return DrawingsExt;
}(Tables);

module.exports = DrawingsExt;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"../drawings":117,"./tables":144,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],137:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var MergedCells = require('./mergedCells');
var toXMLString = require('../XMLString');

var Hyperlinks = function (_MergedCells) {
	(0, _inherits3.default)(Hyperlinks, _MergedCells);

	function Hyperlinks() {
		(0, _classCallCheck3.default)(this, Hyperlinks);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Hyperlinks.__proto__ || (0, _getPrototypeOf2.default)(Hyperlinks)).call(this));

		_this._hyperlinks = [];
		return _this;
	}

	(0, _createClass3.default)(Hyperlinks, [{
		key: 'setHyperlink',
		value: function setHyperlink(hyperlink) {
			hyperlink.objectId = this.common.uniqueId('hyperlink');
			this.relations.addRelation({
				objectId: hyperlink.objectId,
				target: hyperlink.location,
				targetMode: hyperlink.targetMode || 'External'
			}, 'hyperlink');
			this._hyperlinks.push(hyperlink);
			return this;
		}
	}, {
		key: '_insertHyperlink',
		value: function _insertHyperlink(colIndex, rowIndex, hyperlink) {
			var location = void 0;
			var targetMode = void 0;
			var tooltip = void 0;

			if (typeof hyperlink === 'string') {
				location = hyperlink;
			} else {
				location = hyperlink.location;
				targetMode = hyperlink.targetMode;
				tooltip = hyperlink.tooltip;
			}
			this.setHyperlink({
				cell: { c: colIndex + 1, r: rowIndex + 1 },
				location: location,
				targetMode: targetMode,
				tooltip: tooltip
			});
		}
	}, {
		key: '_saveHyperlinks',
		value: function _saveHyperlinks() {
			var relations = this.relations;

			if (this._hyperlinks.length > 0) {
				var children = _.map(this._hyperlinks, function (hyperlink) {
					var attributes = [['ref', util.canonCell(hyperlink.cell)], ['r:id', relations.getRelationshipId(hyperlink)]];

					if (hyperlink.tooltip) {
						attributes.push(['tooltip', hyperlink.tooltip]);
					}
					return toXMLString({
						name: 'hyperlink',
						attributes: attributes
					});
				});

				return toXMLString({
					name: 'hyperlinks',
					children: children
				});
			}
			return '';
		}
	}]);
	return Hyperlinks;
}(MergedCells);

module.exports = Hyperlinks;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"../util":134,"./mergedCells":139,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],138:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var WorksheetSave = require('./save');
var Relations = require('../relations');

var Worksheet = function (_WorksheetSave) {
	(0, _inherits3.default)(Worksheet, _WorksheetSave);

	function Worksheet(workbook) {
		var config = arguments.length > 1 && arguments[1] !== undefined ? arguments[1] : {};
		(0, _classCallCheck3.default)(this, Worksheet);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Worksheet.__proto__ || (0, _getPrototypeOf2.default)(Worksheet)).call(this, workbook, config));

		_this.workbook = workbook;
		_this.common = workbook.common;

		_this.objectId = _this.common.uniqueId('Worksheet');
		_this.data = [];
		_this.columns = [];
		_this.rows = [];

		_this.name = config.name;
		_this.state = config.state || 'visible';
		_this.timezoneOffset = new Date().getTimezoneOffset() * 60 * 1000;
		_this.relations = new Relations(_this.common);
		return _this;
	}

	(0, _createClass3.default)(Worksheet, [{
		key: 'end',
		value: function end() {
			return this.workbook;
		}
	}, {
		key: 'setActive',
		value: function setActive() {
			this.common.setActiveWorksheet(this);
			return this;
		}
	}, {
		key: 'setVisible',
		value: function setVisible() {
			this.setState('visible');
			return this;
		}
	}, {
		key: 'setHidden',
		value: function setHidden() {
			this.setState('hidden');
			return this;
		}
		/**
   * //http://www.datypic.com/sc/ooxml/t-ssml_ST_SheetState.html
   * @param state - visible | hidden | veryHidden
   */

	}, {
		key: 'setState',
		value: function setState(state) {
			this.state = state;
			return this;
		}
	}, {
		key: 'getState',
		value: function getState() {
			return this.state;
		}
	}, {
		key: 'setRows',
		value: function setRows(startRow, rows) {
			var _this2 = this;

			if (!rows) {
				rows = startRow;
				startRow = 0;
			} else {
				--startRow;
			}
			_.forEach(rows, function (row, i) {
				_this2.rows[startRow + i] = row;
			});
			return this;
		}
	}, {
		key: 'setRow',
		value: function setRow(rowIndex, meta) {
			this.rows[--rowIndex] = meta;
			return this;
		}
		/**
   * http://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.column.aspx
   */

	}, {
		key: 'setColumns',
		value: function setColumns(startColumn, columns) {
			var _this3 = this;

			if (!columns) {
				columns = startColumn;
				startColumn = 0;
			} else {
				--startColumn;
			}
			_.forEach(columns, function (column, i) {
				_this3.columns[startColumn + i] = column;
			});
			return this;
		}
	}, {
		key: 'setColumn',
		value: function setColumn(columnIndex, column) {
			this.columns[--columnIndex] = column;
			return this;
		}
	}, {
		key: 'setData',
		value: function setData(startRow, data) {
			var _this4 = this;

			if (!data) {
				data = startRow;
				startRow = 0;
			} else {
				--startRow;
			}
			_.forEach(data, function (row, i) {
				_this4.data[startRow + i] = row;
			});
			return this;
		}
	}]);
	return Worksheet;
}(WorksheetSave);

module.exports = Worksheet;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../relations":120,"./save":142,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],139:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var DrawingsExt = require('./drawing');
var toXMLString = require('../XMLString');

var MergedCells = function (_DrawingsExt) {
	(0, _inherits3.default)(MergedCells, _DrawingsExt);

	function MergedCells() {
		(0, _classCallCheck3.default)(this, MergedCells);

		var _this = (0, _possibleConstructorReturn3.default)(this, (MergedCells.__proto__ || (0, _getPrototypeOf2.default)(MergedCells)).call(this));

		_this._mergedCells = [];
		return _this;
	}

	(0, _createClass3.default)(MergedCells, [{
		key: 'mergeCells',
		value: function mergeCells(cell1, cell2) {
			this._mergedCells.push([cell1, cell2]);
			return this;
		}
	}, {
		key: '_insertMergeCells',
		value: function _insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan) {
			var i = void 0,
			    j = void 0;
			var row = void 0;

			if (colSpan) {
				for (j = 0; j < colSpan; j++) {
					dataRow.splice(colIndex + 1, 0, { style: null, type: 'empty' });
				}
			}
			if (rowSpan) {
				colSpan += 1;

				for (i = 0; i < rowSpan; i++) {
					//todo: original data changed
					row = this.data[rowIndex + i + 1];

					if (!row) {
						row = [];
						this.data[rowIndex + i + 1] = row;
					}

					if (row.length > colIndex) {
						for (j = 0; j < colSpan; j++) {
							row.splice(colIndex, 0, { style: null, type: 'empty' });
						}
					} else {
						for (j = 0; j < colSpan; j++) {
							row[colIndex + j] = { style: null, type: 'empty' };
						}
					}
				}
			}
		}
	}, {
		key: '_saveMergeCells',
		value: function _saveMergeCells() {
			if (this._mergedCells.length > 0) {
				var children = _.map(this._mergedCells, function (mergeCell) {
					return toXMLString({
						name: 'mergeCell',
						attributes: [['ref', util.canonCell(mergeCell[0]) + ':' + util.canonCell(mergeCell[1])]]
					});
				});

				return toXMLString({
					name: 'mergeCells',
					attributes: [['count', this._mergedCells.length]],
					children: children
				});
			}
			return '';
		}
	}]);
	return MergedCells;
}(DrawingsExt);

module.exports = MergedCells;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"../util":134,"./drawing":136,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],140:[function(require,module,exports){
(function (global){
'use strict';

var _typeof2 = require('babel-runtime/helpers/typeof');

var _typeof3 = _interopRequireDefault(_typeof2);

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var SheetView = require('./sheetView');

var PrepareSave = function (_SheetView) {
	(0, _inherits3.default)(PrepareSave, _SheetView);

	function PrepareSave() {
		(0, _classCallCheck3.default)(this, PrepareSave);
		return (0, _possibleConstructorReturn3.default)(this, (PrepareSave.__proto__ || (0, _getPrototypeOf2.default)(PrepareSave)).apply(this, arguments));
	}

	(0, _createClass3.default)(PrepareSave, [{
		key: '_prepare',
		value: function _prepare() {
			var maxX = 0;
			var preparedDataRow = void 0;

			this.preparedData = [];
			this.preparedColumns = [];
			this.preparedRows = [];

			this._prepareTables();
			prepareColumns(this);
			prepareRows(this);

			for (var rowIndex = 0, len = this.data.length; rowIndex < len; rowIndex++) {
				preparedDataRow = prepareDataRow(this, rowIndex);

				if (preparedDataRow.length > maxX) {
					maxX = preparedDataRow.length;
				}
			}

			this.data = null;
			this.columns = null;
			this.rows = null;

			this.maxX = maxX;
			this.maxY = this.preparedData.length;
		}
	}]);
	return PrepareSave;
}(SheetView);

function prepareColumns(worksheet) {
	var styles = worksheet.common.styles;

	_.forEach(worksheet.columns, function (column, index) {
		var preparedColumn = void 0;

		if (column) {
			preparedColumn = _.clone(column);

			if (column.style) {
				preparedColumn.style = styles.addFormat(column.style);
				if (styles._get(preparedColumn.style).fillOut) {
					preparedColumn.styleId = styles._getId(preparedColumn.style);
				} else {
					preparedColumn.styleId = styles._getId(styles._addInvisibleFormat(preparedColumn.style));
				}
			}
			worksheet.preparedColumns[index] = preparedColumn;
		}
	});
}

function prepareRows(worksheet) {
	var styles = worksheet.common.styles;

	_.forEach(worksheet.rows, function (row, index) {
		var preparedRow = void 0;

		if (row) {
			preparedRow = _.clone(row);

			if (row.style) {
				preparedRow.style = styles.addFormat(row.style);
			}
			worksheet.preparedRows[index] = preparedRow;
		}
	});
}

function prepareDataRow(worksheet, rowIndex) {
	var common = worksheet.common;
	var styles = common.styles;
	var preparedDataRow = [];
	var row = worksheet.preparedRows[rowIndex];
	var dataRow = worksheet.data[rowIndex];
	var rowStyle = null;

	if (dataRow) {
		if (!_.isArray(dataRow)) {
			row = mergeDataRowToRow(worksheet, row, dataRow);
			dataRow = dataRow.data;
		}
		if (row) {
			rowStyle = row.style || null;
		}

		for (var colIndex = 0; colIndex < dataRow.length; colIndex++) {
			var column = worksheet.preparedColumns[colIndex];
			var value = dataRow[colIndex];
			var cellStyle = null;
			var cellFormula = null;
			var isString = false;
			var cellValue = void 0;
			var cellType = void 0;

			if (_.isDate(value)) {
				cellValue = value;
				cellType = 'date';
			} else if (value && (typeof value === 'undefined' ? 'undefined' : (0, _typeof3.default)(value)) === 'object') {
				if (value.style) {
					cellStyle = value.style;
				}
				if (value.formula) {
					cellValue = value.formula;
					cellType = 'formula';
				} else if (value.date) {
					cellValue = value.date;
					cellType = 'date';
				} else if (value.time) {
					cellValue = value.time;
					cellType = 'time';
				} else {
					cellValue = value.value;
					cellType = value.type;
				}

				insertEmbedded(worksheet, dataRow, value, colIndex, rowIndex);
			} else {
				cellValue = value;
				cellType = null;
			}

			cellStyle = styles._merge(column ? column.style : null, rowStyle, cellStyle);

			if (!cellType) {
				if (column && column.type) {
					cellType = column.type;
				} else if (typeof cellValue === 'number') {
					cellType = 'number';
				} else if (typeof cellValue === 'string') {
					cellType = 'string';
				}
			}

			if (cellType === 'string') {
				cellValue = common.addString(cellValue);
				isString = true;
			} else if (cellType === 'date' || cellType === 'time') {
				var date = 25569.0 + ((_.isDate(cellValue) ? cellValue.valueOf() : cellValue) - worksheet.timezoneOffset) / (60 * 60 * 24 * 1000);
				if (_.isFinite(date)) {
					cellValue = date;
				} else {
					cellValue = common.addString(String(cellValue));
					isString = true;
				}
			} else if (cellType === 'formula') {
				cellFormula = _.escape(cellValue);
				cellValue = null;
			}

			preparedDataRow[colIndex] = {
				value: cellValue,
				formula: cellFormula,
				styleId: styles._getId(cellStyle),
				isString: isString
			};
		}
	}

	worksheet.preparedData[rowIndex] = preparedDataRow;
	if (row) {
		if (row.style) {
			if (styles._get(row.style).fillOut) {
				row.styleId = styles._getId(row.style);
			} else {
				row.styleId = styles._getId(styles._addInvisibleFormat(row.style));
			}
		}
		worksheet.preparedRows[rowIndex] = row;
	}

	return preparedDataRow;
}

function mergeDataRowToRow(worksheet, row, dataRow) {
	row = row || {};
	row.height = dataRow.height || row.height;
	row.style = dataRow.style ? worksheet.common.styles.addFormat(dataRow.style) : row.style;
	row.outlineLevel = dataRow.outlineLevel || row.outlineLevel;

	return row;
}

function insertEmbedded(worksheet, dataRow, value, colIndex, rowIndex) {
	if (value.hyperlink) {
		worksheet._insertHyperlink(colIndex, rowIndex, value.hyperlink);
	}
	if (value.image) {
		worksheet._insertDrawing(colIndex, rowIndex, value.image);
	}
	if (value.colspan || value.rowspan) {
		var colSpan = (value.colspan || 1) - 1;
		var rowSpan = (value.rowspan || 1) - 1;

		worksheet.mergeCells({ c: colIndex + 1, r: rowIndex + 1 }, { c: colIndex + 1 + colSpan, r: rowIndex + 1 + rowSpan });
		worksheet._insertMergeCells(dataRow, colIndex, rowIndex, colSpan, rowSpan);
	}
}

module.exports = PrepareSave;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"./sheetView":143,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12,"babel-runtime/helpers/typeof":13}],141:[function(require,module,exports){
(function (global){
'use strict';

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var toXMLString = require('../XMLString');

var Print = function () {
	function Print() {
		(0, _classCallCheck3.default)(this, Print);

		this._headers = [];
		this._footers = [];
	}
	/**
  * Expects an array length of three.
  * @param {Array} headers [left, center, right]
  */


	(0, _createClass3.default)(Print, [{
		key: 'setHeader',
		value: function setHeader(headers) {
			if (!_.isArray(headers)) {
				throw 'Invalid argument type - setHeader expects an array of three instructions';
			}
			this._headers = headers;
			return this;
		}
		/**
   * Expects an array length of three.
   * @param {Array} footers [left, center, right]
   */

	}, {
		key: 'setFooter',
		value: function setFooter(footers) {
			if (!_.isArray(footers)) {
				throw 'Invalid argument type - setFooter expects an array of three instructions';
			}
			this._footers = footers;
			return this;
		}
		/**
   * Set page details in inches.
   */

	}, {
		key: 'setPageMargin',
		value: function setPageMargin(margin) {
			this._margin = _.defaults(margin, {
				left: 0.7,
				right: 0.7,
				top: 0.75,
				bottom: 0.75,
				header: 0.3,
				footer: 0.3
			});
			return this;
		}
		/**
   * http://www.datypic.com/sc/ooxml/t-ssml_ST_Orientation.html
   *
   * Can be one of 'portrait' or 'landscape'.
   *
   * @param {String} orientation
   */

	}, {
		key: 'setPageOrientation',
		value: function setPageOrientation(orientation) {
			this._orientation = orientation;
			return this;
		}
		/**
   * Set rows to repeat for print
   *
   * @param {int|[int, int]} params - number of rows to repeat from the top | [first, last] repeat rows
   */

	}, {
		key: 'setPrintTitleTop',
		value: function setPrintTitleTop(params) {
			this._printTitles = this._printTitles || {};

			if (_.isObject(params)) {
				this._printTitles.topFrom = params[0];
				this._printTitles.topTo = params[1];
			} else {
				this._printTitles.topFrom = 0;
				this._printTitles.topTo = params - 1;
			}
			return this;
		}
		/**
   * Set columns to repeat for print
   *
   * @param {int|[int, int]} params - number of columns to repeat from the left | [first, last] repeat columns
   */

	}, {
		key: 'setPrintTitleLeft',
		value: function setPrintTitleLeft(params) {
			this._printTitles = this._printTitles || {};

			if (_.isObject(params)) {
				this._printTitles.leftFrom = params[0];
				this._printTitles.leftTo = params[1];
			} else {
				this._printTitles.leftFrom = 0;
				this._printTitles.leftTo = params - 1;
			}
			return this;
		}
	}, {
		key: '_savePrint',
		value: function _savePrint() {
			return savePageMargins(this._margin) + savePageSetup(this._orientation) + saveHeaderFooter(this._headers, this._footers);
		}
	}]);
	return Print;
}();

function savePageMargins(margin) {
	if (margin) {
		return toXMLString({
			name: 'pageMargins',
			attributes: [['top', margin.top], ['bottom', margin.bottom], ['left', margin.left], ['right', margin.right], ['header', margin.header], ['footer', margin.footer]]
		});
	}
	return '';
}

function savePageSetup(orientation) {
	if (orientation) {
		return toXMLString({
			name: 'pageSetup',
			attributes: [['orientation', orientation]]
		});
	}
	return '';
}

function saveHeaderFooter(headers, footers) {
	if (headers.length > 0 || footers.length > 0) {
		var children = [];

		if (headers.length > 0) {
			children.push(saveHeader(headers));
		}
		if (footers.length > 0) {
			children.push(saveFooter(footers));
		}

		return toXMLString({
			name: 'headerFooter',
			children: children
		});
	}
	return '';
}

function saveHeader(headers) {
	return toXMLString({
		name: 'oddHeader',
		value: compilePageDetailPackage(headers)
	});
}

function saveFooter(footers) {
	return toXMLString({
		name: 'oddFooter',
		value: compilePageDetailPackage(footers)
	});
}

function compilePageDetailPackage(data) {
	data = data || '';

	return ['&L', compilePageDetailPiece(data[0] || ''), '&C', compilePageDetailPiece(data[1] || ''), '&R', compilePageDetailPiece(data[2] || '')].join('');
}

function compilePageDetailPiece(data) {
	if (_.isString(data)) {
		return '&"-,Regular"'.concat(data);
	} else if (_.isObject(data) && !_.isArray(data)) {
		var string = '';

		if (data.font || data.bold) {
			var weighting = data.bold ? 'Bold' : 'Regular';

			string += '&"' + (data.font || '-') + ',' + weighting + '"';
		} else {
			string += '&"-,Regular"';
		}
		if (data.underline) {
			string += '&U';
		}
		if (data.fontSize) {
			string += '&' + data.fontSize;
		}
		string += data.text;

		return string;
	} else if (_.isArray(data)) {
		return _.reduce(data, function (result, value) {
			return result.concat(compilePageDetailPiece(value));
		}, '');
	}
}

module.exports = Print;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10}],142:[function(require,module,exports){
(function (global){
/*eslint global-require: "off"*/

'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var Readable = require('stream').Readable || null;
var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var PrepareSave = require('./prepareSave');
var util = require('../util');
var toXMLString = require('../XMLString');

var WorksheetSave = function (_PrepareSave) {
	(0, _inherits3.default)(WorksheetSave, _PrepareSave);

	function WorksheetSave() {
		(0, _classCallCheck3.default)(this, WorksheetSave);
		return (0, _possibleConstructorReturn3.default)(this, (WorksheetSave.__proto__ || (0, _getPrototypeOf2.default)(WorksheetSave)).apply(this, arguments));
	}

	(0, _createClass3.default)(WorksheetSave, [{
		key: '_save',
		value: function _save(canStream) {
			if (canStream) {
				return new WorksheetStream({
					worksheet: this
				});
			} else {
				return saveBeforeRows(this) + saveData(this) + saveAfterRows(this);
			}
		}
	}]);
	return WorksheetSave;
}(PrepareSave);

var WorksheetStream = function (_Readable) {
	(0, _inherits3.default)(WorksheetStream, _Readable);

	function WorksheetStream(options) {
		(0, _classCallCheck3.default)(this, WorksheetStream);

		var _this2 = (0, _possibleConstructorReturn3.default)(this, (WorksheetStream.__proto__ || (0, _getPrototypeOf2.default)(WorksheetStream)).call(this, options));

		_this2.worksheet = options.worksheet;
		_this2.status = 0;
		return _this2;
	}

	(0, _createClass3.default)(WorksheetStream, [{
		key: '_read',
		value: function _read(size) {
			var worksheet = this.worksheet;
			var stop = false;

			if (this.status === 0) {
				stop = !this.push(saveBeforeRows(worksheet));

				this.status = 1;
				this.index = 0;
				this.len = worksheet.preparedData.length;
			}

			if (this.status === 1) {
				var data = worksheet.preparedData;
				var preparedRows = worksheet.preparedRows;
				var s = '';

				while (this.index < this.len && !stop) {
					while (this.index < this.len && s.length < size) {
						s += saveRow(data[this.index], preparedRows[this.index], this.index);
						data[this.index] = null;
						this.index++;
					}
					stop = !this.push(s);
					s = '';
				}

				if (this.index === this.len) {
					this.status = 2;
				}
			}

			if (this.status === 2) {
				this.push(saveAfterRows(worksheet));
				this.push(null);
			}
		}
	}]);
	return WorksheetStream;
}(Readable);

function saveBeforeRows(worksheet) {
	return util.xmlPrefix + '<worksheet xmlns="' + util.schemas.spreadsheetml + '" xmlns:r="' + util.schemas.relationships + '" xmlns:mc="' + util.schemas.markupCompat + '">' + saveDimension(worksheet.maxX, worksheet.maxY) + worksheet._saveSheetView() + saveColumns(worksheet.preparedColumns) + '<sheetData>';
}

function saveAfterRows(worksheet) {
	return '</sheetData>' +
	// 'mergeCells' should be written before 'headerFoot' and 'drawing' due to issue
	// with Microsoft Excel (2007, 2013)
	worksheet._saveMergeCells() + worksheet._saveHyperlinks() + worksheet._savePrint() + worksheet._saveTables() +
	// the 'drawing' element should be written last, after 'headerFooter', 'mergeCells', etc. due
	// to issue with Microsoft Excel (2007, 2013)
	worksheet._saveDrawing() + '</worksheet>';
}

function saveData(worksheet) {
	var data = worksheet.preparedData;
	var preparedRows = worksheet.preparedRows;
	var children = '';

	for (var i = 0, len = data.length; i < len; i++) {
		children += saveRow(data[i], preparedRows[i], i);
		data[i] = null;
	}
	return children;
}

function saveRow(dataRow, row, rowIndex) {
	var rowLen = void 0;
	var rowChildren = [];
	var colIndex = void 0;
	var value = void 0;
	var attrs = void 0;

	if (dataRow) {
		rowLen = dataRow.length;
		rowChildren = new Array(rowLen);

		for (colIndex = 0; colIndex < rowLen; colIndex++) {
			value = dataRow[colIndex];

			if (!value) {
				continue;
			}

			attrs = ' r="' + util.positionToLetter(colIndex + 1, rowIndex + 1) + '"';
			if (value.styleId) {
				attrs += ' s="' + value.styleId + '"';
			}
			if (value.isString) {
				attrs += ' t="s"';
			}

			if (!_.isNil(value.value)) {
				rowChildren[colIndex] = '<c' + attrs + '><v>' + value.value + '</v></c>';
			} else if (!_.isNil(value.formula)) {
				rowChildren[colIndex] = '<c' + attrs + '><f>' + value.formula + '</f></c>';
			} else {
				rowChildren[colIndex] = '<c' + attrs + '/>';
			}
		}
	}

	return '<row' + getRowAttributes(row, rowIndex) + '>' + rowChildren.join('') + '</row>';
}

function getRowAttributes(row, rowIndex) {
	var attributes = ' r="' + (rowIndex + 1) + '"';

	if (row) {
		if (row.height !== undefined) {
			attributes += ' customHeight="1"';
			attributes += ' ht="' + row.height + '"';
		}
		if (row.styleId) {
			attributes += ' customFormat="1"';
			attributes += ' s="' + row.styleId + '"';
		}
		if (row.outlineLevel) {
			attributes += ' outlineLevel="' + row.outlineLevel + '"';
		}
	}
	return attributes;
}

function saveDimension(maxX, maxY) {
	var attributes = [];

	if (maxX !== 0) {
		attributes.push(['ref', 'A1:' + util.positionToLetter(maxX, maxY)]);
	} else {
		attributes.push(['ref', 'A1']);
	}

	return toXMLString({
		name: 'dimension',
		attributes: attributes
	});
}

function saveColumns(columns) {
	if (columns.length) {
		var children = _.map(columns, function (column, index) {
			column = column || {};

			var attributes = [['min', column.min || index + 1], ['max', column.max || index + 1]];

			if (column.hidden) {
				attributes.push(['hidden', 1]);
			}
			if (column.bestFit) {
				attributes.push(['bestFit', 1]);
			}
			if (column.customWidth || column.width) {
				attributes.push(['customWidth', 1]);
			}
			if (column.width) {
				attributes.push(['width', column.width]);
			} else {
				attributes.push(['width', 9.140625]);
			}
			if (column.styleId) {
				attributes.push(['style', column.styleId]);
			}

			return toXMLString({
				name: 'col',
				attributes: attributes
			});
		});

		return toXMLString({
			name: 'cols',
			children: children
		});
	}
	return '';
}

module.exports = WorksheetSave;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"../util":134,"./prepareSave":140,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12,"stream":108}],143:[function(require,module,exports){
(function (global){
'use strict';

/**
 * https://msdn.microsoft.com/en-us/library/documentformat.openxml.spreadsheet.sheetview.aspx
 */

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var util = require('../util');
var Hyperlinks = require('./hyperlinks');
var toXMLString = require('../XMLString');

var SheetView = function (_Hyperlinks) {
	(0, _inherits3.default)(SheetView, _Hyperlinks);

	function SheetView(workbook, config) {
		(0, _classCallCheck3.default)(this, SheetView);

		var _this = (0, _possibleConstructorReturn3.default)(this, (SheetView.__proto__ || (0, _getPrototypeOf2.default)(SheetView)).call(this, workbook, config));

		_this._pane = null;
		_this._attributes = {
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
				value: null //A1
			},
			view: {
				value: 'normal' //normal | pageBreakPreview | pageLayout
			},
			windowProtection: {
				value: null,
				bool: true
			},
			zoomScale: {
				value: null //10-400
			},
			zoomScaleNormal: {
				value: null //10-400
			},
			zoomScalePageLayoutView: {
				value: null //10-400
			},
			zoomScaleSheetLayoutView: {
				value: null //10-400
			}
		};

		_.forEach(config, function (value, name) {
			if (name === 'freeze') {
				_this.freeze(value.col, value.row, value.cell, value.activePane);
			} else if (name === 'split') {
				_this.split(value.x, value.y, value.cell, value.activePane);
			} else if (_this._attributes[name]) {
				_this._attributes[name].value = value;
			}
		});
		return _this;
	}

	(0, _createClass3.default)(SheetView, [{
		key: 'setAttribute',
		value: function setAttribute(name, value) {
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

	}, {
		key: 'freeze',
		value: function freeze(col, row, cell, activePane) {
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

	}, {
		key: 'split',
		value: function split(x, y, cell, activePane) {
			this._pane = {
				state: 'split',
				xSplit: x * 20,
				ySplit: y * 20,
				topLeftCell: util.canonCell(cell) || 'A1',
				activePane: activePane || 'bottomRight'
			};
			return this;
		}
	}, {
		key: '_saveSheetView',
		value: function _saveSheetView() {
			var attributes = [['workbookViewId', 0]];

			_.forEach(this._attributes, function (attr, name) {
				var value = attr.value;

				if (value !== null) {
					if (attr.bool) {
						value = value ? 'true' : 'false';
					}
					attributes.push([name, value]);
				}
			});

			return toXMLString({
				name: 'sheetViews',
				children: [toXMLString({
					name: 'sheetView',
					attributes: attributes,
					children: [savePane(this._pane)]
				})]
			});
		}
	}]);
	return SheetView;
}(Hyperlinks);

function savePane(pane) {
	if (pane) {
		return toXMLString({
			name: 'pane',
			attributes: [['state', pane.state], ['xSplit', pane.xSplit], ['ySplit', pane.ySplit], ['topLeftCell', pane.topLeftCell], ['activePane', pane.activePane]]
		});
	}
	return '';
}

module.exports = SheetView;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"../util":134,"./hyperlinks":137,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}],144:[function(require,module,exports){
(function (global){
'use strict';

var _getPrototypeOf = require('babel-runtime/core-js/object/get-prototype-of');

var _getPrototypeOf2 = _interopRequireDefault(_getPrototypeOf);

var _classCallCheck2 = require('babel-runtime/helpers/classCallCheck');

var _classCallCheck3 = _interopRequireDefault(_classCallCheck2);

var _createClass2 = require('babel-runtime/helpers/createClass');

var _createClass3 = _interopRequireDefault(_createClass2);

var _possibleConstructorReturn2 = require('babel-runtime/helpers/possibleConstructorReturn');

var _possibleConstructorReturn3 = _interopRequireDefault(_possibleConstructorReturn2);

var _inherits2 = require('babel-runtime/helpers/inherits');

var _inherits3 = _interopRequireDefault(_inherits2);

function _interopRequireDefault(obj) { return obj && obj.__esModule ? obj : { default: obj }; }

var _ = typeof window !== "undefined" ? window['_'] : typeof global !== "undefined" ? global['_'] : null;
var Print = require('./print');
var Table = require('../table');
var toXMLString = require('../XMLString');

var Tables = function (_Print) {
	(0, _inherits3.default)(Tables, _Print);

	function Tables() {
		(0, _classCallCheck3.default)(this, Tables);

		var _this = (0, _possibleConstructorReturn3.default)(this, (Tables.__proto__ || (0, _getPrototypeOf2.default)(Tables)).call(this));

		_this._tables = [];
		return _this;
	}

	(0, _createClass3.default)(Tables, [{
		key: 'addTable',
		value: function addTable(config) {
			var table = new Table(this, config);

			this.common.addTable(table);
			this.relations.addRelation(table, 'table');
			this._tables.push(table);

			return table;
		}
	}, {
		key: '_prepareTables',
		value: function _prepareTables() {
			var data = this.data;

			_.forEach(this._tables, function (table) {
				table._prepare(data);
			});
		}
	}, {
		key: '_saveTables',
		value: function _saveTables() {
			var relations = this.relations;

			if (this._tables.length > 0) {
				var children = _.map(this._tables, function (table) {
					return toXMLString({
						name: 'tablePart',
						attributes: [['r:id', relations.getRelationshipId(table)]]
					});
				});

				return toXMLString({
					name: 'tableParts',
					attributes: [['count', this._tables.length]],
					children: children
				});
			}
			return '';
		}
	}]);
	return Tables;
}(Print);

module.exports = Tables;

}).call(this,typeof global !== "undefined" ? global : typeof self !== "undefined" ? self : typeof window !== "undefined" ? window : {})
},{"../XMLString":109,"../table":133,"./print":141,"babel-runtime/core-js/object/get-prototype-of":5,"babel-runtime/helpers/classCallCheck":9,"babel-runtime/helpers/createClass":10,"babel-runtime/helpers/inherits":11,"babel-runtime/helpers/possibleConstructorReturn":12}]},{},[119])(119)
});