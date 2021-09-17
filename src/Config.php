<?php namespace SamagTech\ExcelLib;

/**
 * Configurazione per i font ed i colori
 *
 */
class Config {

    /**
     * Configurazione per i valori negativi
     *
     * @var array
     */
    public array $negativeStyle = [
        'font' => [
            'color' => [
                'rgb' => 'FF0000'
            ]
        ]
    ];

    /**
     * Configurazione per le righe pari
     *
     * @var string
     */
    public string $rowEven = 'FFFFFF';

    /**
     * Configurazione per le righe dispari
     *
     * @var string
     */
    public string $rowOdd = 'EEEEEE';
}
