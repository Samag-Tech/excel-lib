<?php namespace SamagTech\ExcelLib;

/**
 * Classe per la creazione delle istanze delle libreria
 *
 */
class Factory {

    //------------------------------------------------------------------------------

    /**
     * Crea l'instanza per generare l'excel
     *
     * @param string        $path                   Path dove verrà generato l'excel
     * @param string|null   $filename               Nome del file alla generazione (Opzionale)
     * @param array|null    $ignoreFieldsFormat     Lista dei campi da ignorare (Opzionale)
     *
     * @return Writer
     */
    public function createWriter(string $path, ?string $filename = null, ?array $ignoreFieldsFormat = null, string $config = Config::class, ?string $formatCell = null) : Writer {
        return new Writer($path, $filename, $ignoreFieldsFormat, $config, $formatCell);
    }

    //------------------------------------------------------------------------------

    /**
     * Crea l'instanza leggere dall'excel i dati
     *
     * @param string        $path                   Path dove è situato il file
     * @param string        $filename               Nome del file da caricare
     *
     * @return Reader
     */
    public function createReader(string $path, string $filename) : Reader {
        return new Reader($path, $filename);
    }
}