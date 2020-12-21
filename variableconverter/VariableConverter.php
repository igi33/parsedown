<?php
class VariableConverter {
     // key = var name, value = associative array formatted as:
     // ('source' => function, 'properties' => optional object properties)
    protected $keys;
     // key = var name, value = evaluated value
    protected $values;

    public function __construct() {
        $this->keys = [];
        $this->values = [];
    }

    public function getKeys() {
        return array_keys($this->keys);
    }

    public function getValues() {
        return $this->values;
    }

    public function registerVariable($key, $fnSource, array $properties = []) {
        $this->keys[$key] = ['collection' => false, 'fn_source' => $fnSource, 'properties' => $properties, 'fns_params' => []];
    }

    public function registerCollectionVariable($key, $fnSource, $fnHandle, array $fnsParams = []) {
        $this->keys[$key] = ['collection' => true, 'fn_source' => $fnSource, 'properties' => [], 'fn_handle' => $fnHandle, 'fns_params' => $fnsParams];
    }

    public function unregisterVariable() {
        unset($this->keys[$key]);
        unset($this->values[$key]);
    }

    public function evaluate($key, $params = []) {
        $keyPathItems = explode('.', $key);
        $name = array_shift($keyPathItems);

        if (!array_key_exists($name, $this->keys)) {
            throw new Exception("Cannot evaluate, non-existent key '$name'");
        }

        if (!array_key_exists($name, $this->values)) {
            $this->values[$name] = $this->keys[$name]['fn_source']();
        }

        $value = $this->values[$name];

        if ($this->keys[$name]['collection']) {
            // variable is a collection
            foreach ($this->keys[$name]['fns_params'] as $i => $fnParam) {
                if (isset($params[$i])) {
                    $fnParam($value, $params[$i]);
                }
            }
            return $this->keys[$name]['fn_handle']($value);
        }

        // from here on out, variable is not a collection

        // handles variable properties like 'kancelarija.naziv' separated by dots
        if (count($keyPathItems)) {
            $keyPropertiesPath = implode('.', $keyPathItems);
            if (!in_array($keyPropertiesPath, $this->keys[$name]['properties'])) {
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

        return $value;
    }
}
