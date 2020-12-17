<?php
class VariableConverter {
    protected $keys;
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

    public function registerVariable($key, $fnSource) {
        $this->keys[$key] = $fnSource;
    }

    public function unregisterVariable() {
        unset($this->keys[$key]);
        unset($this->values[$key]);
    }

    public function evaluate($key) {
        if (!array_key_exists($key, $this->values)) {
            $this->values[$key] = $this->keys[$key]();
        }
        return $this->values[$key];
    }
}
