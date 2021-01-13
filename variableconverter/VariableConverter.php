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

    public function registerVariable($key, $fnSource, array $fnParams = [], array $properties = []) {
        if (!is_callable($fnSource)) {
            throw new Exception('$fnSource must be a callable');
        }
        if (!($fnSource instanceof Closure)) {
            $fnSource = Closure::fromCallable($fnSource);
        }
        $fnSource = $fnSource->bindTo($this, $this);

        foreach ($fnParams as $i => &$p) {
            if (!is_callable($p)) {
                throw new Exception('$fnParams['.$i.'] must be a callable');
            }
            if (!($p instanceof Closure)) {
                $p = Closure::fromCallable($p);
            }
            $p = $p->bindTo($this, $this);
        }

        $this->keys[$key] = ['collection' => false, 'fn_source' => $fnSource, 'properties' => $properties, 'fns_params' => $fnParams];
    }

    public function registerCollectionVariable($key, $fnSource, $fnHandle, array $fnParams = []) {
        if (!is_callable($fnSource)) {
            throw new Exception('$fnSource must be a callable');
        }
        if (!($fnSource instanceof Closure)) {
            $fnSource = Closure::fromCallable($fnSource);
        }
        $fnSource = $fnSource->bindTo($this, $this);

        foreach ($fnParams as $i => &$p) {
            if (!is_callable($p)) {
                throw new Exception('$fnParams['.$i.'] must be a callable');
            }
            if (!($p instanceof Closure)) {
                $p = Closure::fromCallable($p);
            }
            $p = $p->bindTo($this, $this);
        }

        $this->keys[$key] = ['collection' => true, 'fn_source' => $fnSource, 'properties' => [], 'fn_handle' => $fnHandle, 'fns_params' => $fnParams];
    }

    public function unregisterVariable($key) {
        unset($this->keys[$key]);
        unset($this->values[$key]);
    }

    public function evaluate($key, $params = null) {
        $keyPathItems = explode('.', $key);
        $name = array_shift($keyPathItems);

        if (!array_key_exists($name, $this->keys)) {
            throw new Exception("Cannot evaluate, non-existent key '$name'");
        }

        // calculate initial variable value from source function
        if (!array_key_exists($name, $this->values)) {
            $this->values[$name] = $this->keys[$name]['fn_source']();
        }

        // get value
        $value = $this->values[$name];

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

        // optionally execute variable parameter functions
        if ($params) {
            foreach ($this->keys[$name]['fns_params'] as $i => $fnParam) {
                if (isset($params[$i])) {
                    $value = $fnParam($value, count($params[$i]) == 1 ? $params[$i][0] : $params[$i]);
                }
            }
        }

        // return value for non-collections or handled value for collections
        return $this->keys[$name]['collection'] ? $this->keys[$name]['fn_handle']($value) : $value;
    }
}
