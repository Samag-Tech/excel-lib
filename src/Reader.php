<?php namespace SamagTech\ExcelLib;

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

/**
 * Classe per la lettura dei dati da un file xlsx
 *
 */
class Reader extends AbstractExcel {

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
     * Lista e/o nome del foglio che si vuole caricare
     *
     * @access private
     * @var string|array
     */
    private string|array|null $sheetNames = null;

    /**
     * Flag che indica se è presente l'intestazione
     *
     * @access private
     * @var bool
     * Default true
     */
    private bool $hasHeader = true;


    /**
     * Lista che indica la chiave ad associare alla colonna
     *
     *  @access private
     * @var array<string,string>
     */
    private array $columnToKey = [];

    /**
     * Numero di fogli caricati
     *
     * @var int
     * @access private
     */
    private int $numSheet = 1;

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
     * Setto la lista dei fogli da caricare
     *
     * @param string|array  $sheetNames Lista o nome dei/del file da caricare
     *
     * @return self
     */
    public function setSheetNames(array|string $sheetNames) : self {
        $this->sheetNames = $sheetNames;
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
     * Restituisce il numero di fogli caricati
     *
     * @return int
     */
    public function getNumSheet() : int  { return $this->numSheet; }

    //------------------------------------------------------------------------------

    /**
     * Restituisce i dati dell'excel in formato array
     *
     * @throws ExcelException   Solleva quest'eccezione in caso di errore
     *
     * @return array
     */
    public function toArray() : array {
        return $this->build();

    }

    //------------------------------------------------------------------------------

    /**
     * Restituisce i dati dell'excel in formato object
     *
     * @throws ExcelException   Solleva quest'eccezione in caso di errore
     *
     * @return object
     */
    public function toObject() : object{
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

        // Recupero tutti i fogli
        $worksheets = $this->spreadsheet->getAllSheets();

        // Controllo come devono essere gestiti i fogli
        if ( count($worksheets)  == 1  ) {
            return $this->manageBody(current($worksheets));
        }
        else if ( count($worksheets) > 1 ) {

            $data = [];

            foreach ($worksheets as $worksheet ) {

                $data[str_replace(' ', '_', $worksheet->getTitle())] = $this->manageBody($worksheet);

                ++$this->numSheet;
            }

            return $data;
        }
        else throw new ExcelException('Il file non ha dati');
    }

    //------------------------------------------------------------------------------

    /**
     *  Funzione per la gestione del corpo di un singolo foglio
     *
     * @param Worksheet Foglio corrente da gestire
     *
     * @return array    Ritorna i dati estratti
     */
    private function manageBody ( Worksheet $worksheet ) : array {

        // Trasforma il foglio in array
        $this->body = $worksheet->toArray();

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

        /**
         * Controlo se sono impostati i fogli
         *
         * Se i fogli sono impostati li carico,
         * altrimenti li carico tutti
         *
         */
        if ( is_null($this->sheetNames) ) {
            $this->reader->setLoadAllSheets();
        }
        else {
            $this->reader->setLoadSheetsOnly($this->sheetNames);
        }

        return $this->reader->load($this->path.$this->filename);
    }

    //------------------------------------------------------------------------------
}