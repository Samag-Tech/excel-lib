<?php namespace SamagTech\ExcelLib;

use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class FormatCell {

    /**
     * Foglio su cui modificare i dati
     *
     * @access private
     * @var Worksheet
     */
    private Worksheet $worksheet;

    /**
     * Configurazione per lo stile ed il font
     *
     * @access private
     * @var Config
     */
    private Config $config;

    /**
     * Indice corrente della cella
     *
     * @access private
     * @var string
     */
    private string $index = 'A1';

    //------------------------------------------------------------------------------

    /**
     * Costruttore
     *
     * @param Worksheet $worksheet  Foglio
     * @param Config    $config     Configurazione
     *
     */
    public function __construct(Worksheet $worksheet, Config $config) {
        $this->worksheet = $worksheet;
        $this->config    = $config;
    }

    //------------------------------------------------------------------------------

    /**
     * Setta l'indice della cella
     *
     * @param string index  Indice corrente della cella
     *
     * @return void
     */
    public function setIndex(string $index) : void {
        $this->index = $index;
    }

    //------------------------------------------------------------------------------

    /**
     * Restituisce l'indice corrente della cella
     *
     * @return string
     */
    public function getIndex() : string {
        return $this->index;
    }

    //------------------------------------------------------------------------------

    /**
     * Imposta la cella con il font Bold
     *
     * @return void
     */
    public function setBold() : void {
        $this->worksheet->getStyle($this->index)->getFont()->setBold(true);
    }

    //------------------------------------------------------------------------------

    /**
     * Imposta il background delle celle in base al numero della riga
     * se pari o dispari.
     *
     * Il colore è applicato dallo stile configurato
     *
     * @param int $row  Numero della riga
     *
     * @return void
     */
    public function setBackground(int $row) : void {

        $background = $row % 2 == 0 ? $this->config->rowEven : $this->config->rowOdd;

        $this->worksheet->getStyle($this->index)
            ->getFill()
            ->setFillType(Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB($background);

    }

    //------------------------------------------------------------------------------

    /**
     * Formatta il numero all'italiana.
     *
     * Se il numero è negativo viene applicato lo stile configurato
     *
     * @param int|float $value     Numero che si deve formattare
     *
     * @return void
     */
    public function setNumberFormat(int|float $value) : void {

        $this->worksheet->getStyle($this->index)
            ->getNumberFormat()
            ->setFormatCode(NumberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);

        // Controllo se è il numero è negativo
        if ( $value < 0 ) {
            $this->worksheet->getStyle($this->index)->applyFromArray($this->config->negativeStyle);
        }
    }

    //------------------------------------------------------------------------------

    /**
     * Formatta il numero in percentuale.
     *
     * Se il numero è negativo viene applicato lo stile configurato
     *
     * @param int|float $value     Numero che si deve formattare
     *
     * @return void
     */
    public function setPercentageFormat(int|float $value) : void {

        $this->worksheet->getStyle($this->index)
            ->getNumberFormat()
            ->setFormatCode(NumberFormat::FORMAT_PERCENTAGE_00);

        // Controllo se è il numero è negativo
        if ( $value < 0 ) {
            $this->worksheet->getStyle($this->index)->applyFromArray($this->config->negativeStyle);
        }

    }


    //------------------------------------------------------------------------------
    public function setDateFormat() {}
    //------------------------------------------------------------------------------
    public function setDateTimeFormat() {}

    //------------------------------------------------------------------------------

}