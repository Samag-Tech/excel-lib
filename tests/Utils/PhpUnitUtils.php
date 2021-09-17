<?php

class PhpUnitUtils {

    /**
     * Funzione per invocare metodi privati o protetti
     *
     * @param $object               Classe da richiamare
     * @param string $method        nome del metodo
     * @param array $parameters     Parametri da passare al metodo
     *
     * @return mixed
     * @throws \Exception
     */
    public static function callMethod($object, string $method , array $parameters = []) {
        try {

            $className = get_class($object);
            $reflection = new \ReflectionClass($className);

        } catch (\ReflectionException $e) {
           throw new \Exception($e->getMessage());
        }

        $method = $reflection->getMethod($method);
        $method->setAccessible(true);

        return $method->invokeArgs($object, $parameters);
    }
}