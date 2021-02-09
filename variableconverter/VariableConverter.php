<?php
class VariableConverter {
     // key = var name, value = associative array formatted as:
     // ['fn_source' => function to get initial value,
     // 'properties' => optional allowed object properties,
     // 'fns_params' => array of functions that handles each parameter of a variable,
     // 'collection' => boolean denoting if the variable is a collection
     // 'fn_handle' => function that transforms the collection values to text]
     protected $keys;
     protected $values; // key = var name, value = evaluated value
 
     public function __construct() {
         $this->keys = [];
         $this->values = [];
     }
     
     public function __set($name, $value) {
         $this->$name = $value;
     }
     
     public function __get($name) {
         if (!isset($this->$name)) {
             throw new Exception("Property '$name' doesn't exist!");
         }
         
         return $this->$name;
     }
     
     public function __isset($name) {
         return isset($this->$name);
     }
 
     public function getKeys() {
         return array_keys($this->keys);
     }
 
     public function getValues() {
         return $this->values;
     }
     
     public function exists($key) {
         return array_key_exists($key, $this->keys);
     }
 
     public function registerVariable($key, $fnSource, array $fnsParams = [], array $properties = []) {
         if (!is_callable($fnSource)) {
             throw new Exception('$fnSource must be a callable');
         }
         if (!($fnSource instanceof Closure)) {
              $fnSource = function () use ($fnSource) { return $fnSource(); };
         }
         $fnSource = $fnSource->bindTo($this, $this);
         
         foreach ($fnsParams as $i => &$p) {
             if (!is_callable($p['callable'])) {
                 throw new Exception('$fnParams['.$i.'] must be a callable');
             }
             if (!($p['callable'] instanceof Closure)) {
                 $p['callable'] = function ($value, $args) use ($p) { return $p['callable']($value, $args); };
             }
             $p['callable'] = $p['callable']->bindTo($this, $this);
         }
         
         $this->keys[$key] = ['collection' => false, 'fn_source' => $fnSource, 'properties' => $properties, 'fns_params' => $fnsParams];
     }
 
     public function registerCollectionVariable($key, $fnSource, $fnPrint, array $fnsParams = []) {
         if (!is_callable($fnSource)) {
             throw new Exception('$fnSource must be a callable');
         }
         if (!($fnSource instanceof Closure)) {
             $fnSource = function () use ($fnSource) { return $fnSource(); };
         }
         $fnSource = $fnSource->bindTo($this, $this);
         
         foreach ($fnsParams as $i => &$p) {
             if (!is_callable($p['callable'])) {
                 throw new Exception('$fnParams['.$i.'] must be a callable');
             }
             if (!($p['callable'] instanceof Closure)) {
                 $p['callable'] = function ($value, $args) use ($p) { return $p['callable']($value, $args); };
             }
             $p['callable'] = $p['callable']->bindTo($this, $this);
         }
         
         if (!is_callable($fnPrint)) {
             throw new Exception('$fnPrint must be a callable');
         }
         if (!($fnPrint instanceof Closure)) {
             $fnPrint = function ($items) use ($fnPrint) { return $fnPrint($items); };
         }
         $fnPrint = $fnPrint->bindTo($this, $this);
         
         $this->keys[$key] = ['collection' => true, 'fn_source' => $fnSource, 'properties' => [], 'fn_handle' => $fnPrint, 'fns_params' => $fnsParams];
     }
 
     public function unregisterVariable($key) {
         unset($this->keys[$key]);
         unset($this->values[$key]);
     }
     
     // Sets a print function for an existing variable.
     // This works for non-collection variables too!
     public function setPrintFn($key, $fnPrint) {
         if (!array_key_exists($key, $this->keys)) {
             throw new Exception("Non-existent key '$key'");
         }
         if (!is_callable($fnPrint)) {
             throw new Exception('$fnPrint must be a callable');
         }
         if (!($fnPrint instanceof Closure)) {
             $fnPrint = function ($items) use ($fnPrint) { return $fnPrint($items); };
         }
         $fnPrint = $fnPrint->bindTo($this, $this);
         
         $this->keys[$key]['fn_handle'] = $fnPrint;
     }
     
     public function getValue($key) {
         if (!array_key_exists($key, $this->keys)) {
             throw new Exception("Cannot get value, non-existent key '$key'");
         }
         
         // calculate value from source function
         if (!array_key_exists($key, $this->values)) {
             $this->values[$key] = $this->keys[$key]['fn_source']();
         }
         
         // get value
         return $this->values[$key];
     }
     
     public function getValueWithParameters($key, $params) {
         $keyPathItems = explode('.', $key);
         $name = array_shift($keyPathItems);
         
         // get value
         $value = $this->getValue($name);
         
         // Handles variable properties like 'kancelarija.naziv' separated by dots.
         // This is useful for cases when a variable value is an associated array or an object.
         // The code checks if the property is an array key, a public object property
         // or if a respective getter object method exists, in that order.
         if (!$this->keys[$name]['collection'] && count($keyPathItems)) {
             $keyPropertiesPath = implode('.', $keyPathItems);
             if ($this->keys[$name]['properties'] && !in_array($keyPropertiesPath, $this->keys[$name]['properties'])) {
                 throw new Exception("Property '$keyPropertiesPath' not in list of allowed properties for key '$name'");
             }
             
             $currentKeyName = $name;
             foreach ($keyPathItems as $prop) {
                 $ake = array_key_exists($prop, $value);
                 $iso = !$ake && is_object($value);
                 $pe = $iso && property_exists($value, $prop);
                 $getterName = 'get'.ucfirst($prop);
                 $me = $iso && method_exists($value, $getterName);
                 
                 if (!$ake && !$pe && !$me) {
                     throw new Exception("Unknown property '$prop' of key '$currentKeyName'");
                 }
                 
                 if ($ake) {
                     $value = $value[$prop];
                 } elseif ($pe) {
                     $value = $value->$prop;
                 } else {
                     $value = $value->{$getterName}();
                 }
                 
                 $currentKeyName .= ".$prop";
             }
         }
         
         // execute variable parameter functions
         foreach ($this->keys[$name]['fns_params'] as $i => $fnParam) {
             $paramValue = $fnParam['default'];
             if (isset($params[$i])) {
                 $paramValue = count($params[$i]) == 1 ? $params[$i][0] : $params[$i];
             }
             $value = $fnParam['callable']($value, $paramValue);
         }
         
         return $value;
     }
 
     public function evaluate($key, $params = null) {
         $keyPathItems = explode('.', $key);
         $name = array_shift($keyPathItems);
         $value = $this->getValueWithParameters($key, $params);
 
         // return print value
         return isset($this->keys[$name]['fn_handle']) ? $this->keys[$name]['fn_handle']($value) : $value;
     }
 }
 