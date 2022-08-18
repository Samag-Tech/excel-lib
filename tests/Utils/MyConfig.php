<?php

use SamagTech\ExcelLib\Config;

class MyConfig extends Config {

    /**
     * Configurazione per le righe pari
     *
     * @var string
     */
    public string $rowEven = '000000';

    /**
     * Configurazione per le righe dispari
     *
     * @var string
     */
    public string $rowOdd = '00FF00';

}
