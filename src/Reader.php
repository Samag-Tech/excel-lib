<?php namespace SamagTech\ExcelLib;

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Classe per la lettura dei dati da un file xlsx
 *
 */
class Reader extends AbstractExcel {

    /**
     * Lista e/o nome del foglio che si vuole caricare
     *
     * @access private
     * @var string|array
     */
    private string|array $sheetNames;

    /**
     * Flag che indica se è presente l'intestazione
     *
     * @access private
     * @var bool
     * Default true
     */
    private bool $hasHeader = true;

    /**
     * Lista dell'intestazione
     *
     * @access protected
     * @var array<string[]>
     */
    protected array $header;

    /**
     * Corpo dei dati dell'excel
     *
     * @access protected
     * @var array[]
     */
    protected array $body;

    /**
     * Lista che indica la chiave ad associare alla colonna
     *
     *  @access private
     * @var array<string,string>
     */
    private array $columnToKey = [];

    //------------------------------------------------------------------------------

    /**
     * {@inheritDoc}
     *
     */
    public function __construct(string $path, string $filename) {
        parent::__construct($path, $filename);

        $this->reader = new Xlsx($this->path.$this->filename);
    }

    //------------------------------------------------------------------------------

    /**
     * Setta l'array per la definizione dell'associazione tra colonne e chiavi dell'array
     *
     * @param array<string,string> $columnToKey
     *
     * @return self
     */
    public function setColumnToKey(array $columnToKey) : self {
        $this->columnToKey = $columnToKey;
        return $this;
    }

    //------------------------------------------------------------------------------

    /**
     * Imposta il flag che indica se è presente l'intestazione o meno
     *
     * @param bool $hasHeader
     *
     * @return self
     */
    public function setHasHeader(bool $hasHeader ) {
        $this->hasHeader = $hasHeader;
        return $this;
    }

    //------------------------------------------------------------------------------

    /**
     * Restituisce i dati dell'excel in formato array
     *
     * @return array
     */
    public function buildArray() : array {
        return $this->build();

    }

    //------------------------------------------------------------------------------

    /**
     * Restituisce i dati dell'excel in formato object
     *
     * @return object
     */
    public function buildObject() : object{
        return json_decode(json_encode($this->build()));
    }

    //------------------------------------------------------------------------------

    /**
     * Esegue la costruzione dell'array dall'excel con le formattazioni
     *
     * @throws ExcelException Solleva quest'eccezione in caso di corpo vuoto
     *
     * @return array
     */
    protected function build () : array  {

        // Carica il file
        $this->spreadsheet = $this->loadFile();

        // Trasforma il foglio in array
        $this->body = $this->spreadsheet->getActiveSheet()->toArray();

        // Se è presente l'intestazione la scorporo dal corpo
        if ( $this->hasHeader ) {
            $this->header = $this->body[0];
            unset($this->body[0]);
        }

        // Se il corpo è vuoto lancia un eccezione
        if ( empty($this->body) ) throw new ExcelException('Il file è vuoto');


        /**
         * Se è presente l'instazione e la definzione delle colonne eseguo il parsing
         * per assegnare la chiave dell'array in base all'intestazione
         */
        if ( ! empty($this->columnToKey) && $this->hasHeader) {
            $this->body = $this->parse();
        }

        return $this->body;
    }

    //------------------------------------------------------------------------------

    /**
     * Esegue il parsing del file associando ad ogni colonna la chiave custom dell'array
     *
     * NB. Le colonne senza definizione vengono scartate
     *
     * @return array
     */
    private function parse() : array {

        $data = [];

        foreach ( $this->body as $cells ) {

            $tmp = [];
            foreach ( $cells as $columnId => $cell) {

                if ( isset($this->columnToKey[$this->header[$columnId]]) ) {
                    $tmp[$this->columnToKey[$this->header[$columnId]]] = $cell;
                }

            }

            $data[] = $tmp;

        }

        return $data;

    }

    //------------------------------------------------------------------------------

    /**
     * Restituisce il file caricato
     *
     * @throws ExcelException   Solleva un eccezione se il file non esiste
     *
     * @return Spreadsheet
     */
    private function loadFile() : Spreadsheet {

        if ( ! $this->filenameExist() ) throw new ExcelException('Il file non esiste');

        $this->reader->setLoadAllSheets();
        return $this->reader->load($this->path.$this->filename);
    }

    //------------------------------------------------------------------------------
}