<?php namespace SamagTech\ExcelLib;

use PhpOffice\PhpSpreadsheet\Reader\Xlsx as ReaderXlsx;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx as WriterXlsx;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

/**
 * Classe per la definizione delle proprietà e funzioni
 * di base per la gestione dell'excel
 *
 * @abstract
 */
abstract class AbstractExcel {

    /**
     * Istanza del foglio
     *
     * @access protected
     * @var Spredsheet
     */
    protected Spreadsheet $spreadsheet;

    /**
     * Istanza per la scrittura del foglio in formato xlsx
     *
     * @access protected
     * @var WriterXlsx
     */
    protected WriterXlsx $writer;

    /**
     * Istanza per la lettura del foglio in formato xlsx
     *
     * @access protected
     * @var ReaderXlsx
     */
    protected ReaderXlsx $reader;

    /**
     * Configurazione per i font e stile dei fogli
     *
     * @access protected
     * @var Config
     */
    protected Config $config;

    /**
     * Nome del file
     *
     * @access protected
     * @var string|null
     */
    protected ?string $filename = 'file';

    /**
     * Path in cui deve essere salvato il file
     *
     * @access protected
     * @var string
     */
    protected string $path;

    /**
     * Lista delle colonne il cui formato deve essere ignorato
     *
     * @access protected
     * @var string[]
     */
    protected array $ignoreFieldsFormat = [];

    /**
     * Estensione dei file xlsx
     *
     * @const
     * @var string
     */
    const EXT = 'xlsx';

    //------------------------------------------------------------------------------

    /**
     * Costruttore.
     *
     * @param string    $path                   Path in cui il file viene salvato
     * @param string    $filename               Nome del file (Opzionale)
     * @param string[]  $ignoreFieldsFormat     Lista delle colonne da ignorare la formattazione (Opzionale)
     * @param string    $config                 Configurazione del file excel (Opzionale)
     *
     * @return void
     */
    public function __construct(string $path, ?string $filename = null, ?array $ignoreFieldsFormat = null, string $config = Config::class) {

        // Inizializzo il path
        $this->setPath($path);

        // Inizializzo i dati opzionali se settati
        if ( ! is_null($filename) ) {
            $this->setFilename($filename);
        }
        else {
            $this->filename = date('Y_m_d__h_i_s_').'file.xlsx';
        }

        if ( ! is_null($ignoreFieldsFormat) ) {
            $this->ignoreFieldsFormat = $ignoreFieldsFormat;
        }

        // Inizializzo l'istanza del foglio
        $this->spreadsheet = new Spreadsheet();

        // Controllo se la configurazione è un istanza di config
        $this->config = new $config instanceof Config ? new $config : throw new ExcelException('La configurazione deve estendere la classe SamagTech\\ExcelLib\\Config');
    }

    //------------------------------------------------------------------------------

    /**
     * Setta il nome del file
     *
     * @param string $filename
     *
     * @return self
     */
    public function setFilename(string $filename) : self {
        $split = explode('.', $filename);
        $this->filename = end($split) == self::EXT ? $filename : $filename.'.'.self::EXT;
        return $this;
    }

    //------------------------------------------------------------------------------

    /**
     * Setta il path del file
     *
     * @param string $path
     *
     * @return self
     */
    public function setPath ( string $path ) {
        $this->path = str_ends_with($path, '/') ? $path : $path.'/';
        return $this;
    }

    //------------------------------------------------------------------------------

    /**
     * Setta la lista delle colonne da ignorare
     *
     * @param string $ignoreFieldsFormat
     *
     * @return self
     */
    public function setIgnoreFieldsFormat ( array $ignoreFieldsFormat ) : self {
        $this->ignoreFieldsFormat = $ignoreFieldsFormat;
        return $this;
    }

    //------------------------------------------------------------------------------

    /**
     * Restituisce il nome del file
     *
     * @return string|null
     */
    public function getFilename() : ?string { return $this->filename; }

    //------------------------------------------------------------------------------

    /**
     * Restituisce il path impostato
     *
     * @return string
     */
    public function getPath() : string { return $this->path; }

    //------------------------------------------------------------------------------

    /**
     * Restituisce il nome del file
     *
     * @return string|null
     */
    public function getIgnoreFieldsFormat() : ?array { return $this->ignoreFieldsFormat; }

    //------------------------------------------------------------------------------

    /**
     * Controlla se il nome del file nel path esiste
     *
     * @return bool     True se il nome esiste, False altrimenti
     */
    protected function filenameExist() : bool {
        return file_exists($this->path.$this->filename);
    }

    //------------------------------------------------------------------------------

    /**
     * Controlla se il path esiste
     *
     * @return bool     True se il path del file esiste, False altrimenti
     */
    protected function pathExist() : bool {
        return file_exists($this->path);
    }

    //------------------------------------------------------------------------------

    /**
     * Crea la cartella
     *
     * @return bool True se la cartella esiste e/o è stata creata, False altrimenti
     */
    protected function createPath() : bool {

        if ( $this->pathExist() ) {
            return true;
        }

        return mkdir($this->path, 0777, true);
    }

    //------------------------------------------------------------------------------
}
