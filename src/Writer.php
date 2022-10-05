<?php namespace SamagTech\ExcelLib;

use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use SamagTech\ExcelLib\AbstractExcel;


class Writer extends AbstractExcel {

    /**
     * Lista dei titoli delle colonne.
     *
     * @access private
     * @var string[]|array<string[]>
     */
    private array $headers = [];

    /**
     * Corpo dei dati dell'excel
     *
     * @access private
     * @var array
     */
    private array $body = [];

    /**
     * Lista delle definizioni dei formati delle colonne
     *
     * Nel caso di chiavi 'date' o 'datetime' aggiungere le chiavi
     * 'old_format' e 'new_format' per specificare il vecchio ed il
     * nuovo formato, di default sono impostati a:
     *  - date : old_format -> Y-m-d h:i:s, new_format -> d/m/Y h:i:s
     *  - datetime : old_format -> Y-m-d h:i:s, new_format -> d/m/Y h:i:s
     *
     * @access private
     * @var array<string, string>
     */
    private array $columnDefinition = [];

    /**
     * Istanza della classe per la formattazione delle celle
     *
     * @access private
     * @var FormatCell
     */
    private FormatCell $formatCell;

    //------------------------------------------------------------------------------

    /**
     * {@inheritDoc}
     *
     */
    public function __construct(string $path, ?string $filename = null, ?array $ignoreFieldsFormat = null, string $config = Config::class, ?string $formatCell = null) {
        parent::__construct($path, $filename, $ignoreFieldsFormat, $config);

        $this->formatCell = ! is_null($formatCell) ?
            new $formatCell($this->spreadsheet->setActiveSheetIndex(0), $this->config)
            : new FormatCell($this->spreadsheet->setActiveSheetIndex(0), $this->config);
    }

    //------------------------------------------------------------------------------

    /**
     * Setta l'intestazione della tabella
     *
     * @param string[]|array<string[]> $headers
     *
     * @return self
     */
    public function setHeader(array $headers) : self {
        $this->headers = $headers;
        return $this;
    }

    //------------------------------------------------------------------------------

    /**
     * Setta la definizione delle colonne
     *
     * @param string[]|array<string[]> $columnDefinition
     *
     * @return self
     */
    public function setColumnDefinition(array $columnDefinition) : self {
        $this->columnDefinition = $columnDefinition;
        return $this;
    }

    //------------------------------------------------------------------------------

    /**
     * Setta il corpo del foglio
     *
     * @param array $body
     *
     * @return self
     */
    public function setBody(array $body) : self {
        $this->body = $body;
        return $this;
    }

    //------------------------------------------------------------------------------

    /**
     * Costruisce il foglio excel e restituisce il path
     *
     * @param bool  $isMultisheet   Flag che indica se è un esportazione multifoglio
     *
     * @return string   Path dove è stato creato il file
     */
    public function build(bool $isMultisheet = false) : string {

        // Controllo se il corpo è presente
        if ( empty($this->body) ) {
            throw new ExcelException('Il corpo dell\'excel è vuoto');
        }

        // Controllo se
        if ( isset($this->body) && ! is_array(current($this->body)) ) {
            throw new ExcelException('Il corpo dell\'excel è malformato');
        }

        // Creo la cartella
        if ( ! $this->createPath() ) {
            throw new ExcelException('Errore della creazione del path');
        }

        if ( ! $isMultisheet ) {

            // Genero i dati del foglio
            $this->generateSheet($this->headers, $this->body);
        }
        else {

            // I fogli sono dati dalle chiavi del corpo
            $sheets = array_keys($this->body);

            // Il numero di fogli è dato dal numero di chiavi del corpo
            $numSheet = count($sheets) - 1;

            // Contatore dei fogli
            $counter = 0;

            // Creo tutti i fogli
            for ( $i = 0; $i < $numSheet ; ++$i ) {
                $this->spreadsheet->createSheet($i);
            }

            // Creo un foglio alla volta
            foreach ( $sheets as $sheet ) {

                // Recupero l'intestazione
                $header = [];

                if ( ! empty($this->headers) && is_array($this->headers[$counter]) ) {
                    $header = $this->headers[$counter];
                }
                else if ( ! empty($this->headers) && is_string(current($this->headers)) ) {
                    $header = $this->headers;
                }

                $this->generateSheet($header, $this->body[$sheet], $numSheet, 'Foglio '. $sheet + 1);

                // Incremento il contatore per l'header
                ++$counter;

                // Decremento il numero di foglio per la costruzione del successivo
                --$numSheet;

            }
        }

        // Salvo il file
        $this->save();

        return $this->path.$this->filename;
    }

    //------------------------------------------------------------------------------

    /**
     * Genera il corpo del foglio
     *
     * @return void
     */
    private function generateSheet(array $header, array $body, int $indexSheet = 0, string $sheetName = 'Foglio 1' ) : void {

        // Imposto la prima colonna e la prima riga
        $column = 'A';
        $row    = 1;

        // Recupero il foglio corrente
        $currentSheet = $this->spreadsheet->setActiveSheetIndex($indexSheet);

        // Imposto il nome dei vari fogli
        $currentSheet->setTitle($sheetName);

        // Istanzio la classe per la formattazione della cella
        $this->formatCell->setWorksheet($currentSheet);

        /**
         * Se esiste l'intestazione viene generata
         * e la riga bloccata.
         *
         * Il contatore dell riga partirà da 2
         */
        if ( ! empty($header) ) {

            $this->generateHeader($currentSheet, $header);

            $currentSheet->freezePane('A2');

            ++$row;
        }

        // Imposto il dimensionamento della cella
        $currentSheet->getDefaultColumnDimension()->setAutoSize(true);

        // Inserisco i dati del corpo
        foreach ( $body as $data ) {

            foreach ( $data as $key => $value ) {

                // Punto l'indice
                $index = $column.$row;

                // Setto l'indice per la formattazione della cella
                $this->formatCell->setIndex($index);

                // setto il valore
                $currentSheet->setCellValue($index, $value);

                // Imposto lo sfondo della cella
                $this->formatCell->setBackground($row);

                // Controllo se devo ignorare altre formattazioni
                if ( !empty($value) && !in_array($key, $this->ignoreFieldsFormat) ) {

                    /**
                     * Se per la colonna è stata settata definizione
                     * eseguo la formattazione altrimenti controllo se è un numero
                     * oppure una stringa
                     *
                     */
                    if ( isset($this->columnDefinition[$key]) ) {

                        switch ($this->columnDefinition[$key] ) {
                            case ExcelEnum::DEFINITION_NUMBER :
                                $this->formatCell->setNumberFormat($value);
                            break;
                            case ExcelEnum::DEFINITION_PERCENTAGE :
                                $this->formatCell->setPercentageFormat($value);
                            break;

                        }
                    }
                    else if ( is_numeric($value) ) {
                        $this->formatCell->setNumberFormat($value);
                    }

                    if ( isset($this->columnDefinition[$key]['type']) ) {

                        switch ($this->columnDefinition[$key]['type'] ) {
                            case ExcelEnum::DEFINITION_DATE:
                                $oldFormat = $this->columnDefinition[$key]['old_format'] ?? 'Y-m-d';
                                $newFormat = $this->columnDefinition[$key]['new_format'] ?? 'd/m/Y';


                                $this->formatCell->setDateFormat($value, $oldFormat, $newFormat);
                            break;
                            case ExcelEnum::DEFINITION_DATETIME :

                                $oldFormat = $this->columnDefinition[$key]['old_format'] ?? 'Y-m-d h:i:s';
                                $newFormat = $this->columnDefinition[$key]['new_format'] ?? 'd/m/Y h:i:s';

                                $this->formatCell->setDateFormat($value, $oldFormat, $newFormat );
                            break;
                        }
                    }
                    else if ( is_numeric($value) ) {
                        $this->formatCell->setNumberFormat($value);
                    }

                }

                // Passo alla colonna successiva
                ++$column;
            }

            // Passo alla riga successiva e imposto la colonna iniziale
            ++$row;
            $column = 'A';

        }
    }

    //------------------------------------------------------------------------------

    /**
     * Funzione che genera l'intestazione del foglio corrente del foglio
     *
     * @param Worksheet $worksheet          Foglio corrente su cui si sta lavorando
     * @param string[]  $currentHeaders     Intestazione corrente
     *
     * @return void
     */
    private function generateHeader(Worksheet $worksheet, array $currentHeaders) {

        // L'intestazione è puntata alla prima riga
        $column = 'A';
        $row    = 1;

        $formatCell = $this->formatCell;

        $formatCell->setWorksheet($worksheet);

        foreach ( $currentHeaders as $header) {

            $index = $column.$row;

            // Imposto il valore nella cella
            $worksheet->setCellValue($index, $header);

            // Imposto il font
            $formatCell->setIndex($index);
            $formatCell->setBold();

            ++$column;

        }
    }

    //------------------------------------------------------------------------------

    /**
     * Funzione che salva il file
     *
     * @return void
     */
    protected function save() : void {
        $this->writer = new Xlsx($this->spreadsheet);
        $this->writer->save($this->path.$this->filename);
    }
}