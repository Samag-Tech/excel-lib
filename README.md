# Libreria per la gestione degli Excel

## Prefazione
Questa libreria è un wrapper per la libreria <a href="https://phpspreadsheet.readthedocs.io/en/latest/">PHPSpreadSheet</a>, una versione avanzata di PHPExcel che non ha problemi con versioni PHP 7.4+ ed è possibile installarla con

```bash
composer require samagtech/excel-lib
```

In caso di modifiche alla libreria e/o aggiunta di funzionalità consigliato lanciare/scrivere gli UnitTest per controllare l'integrità delle funzioni utilizzando il comando

```bash
composer test
```

Per installare PHPUnit (<a href="https://phpunit.readthedocs.io/en/9.5/">docs</a>) lanciare il comando
```bash
composer install --dev
```

Inoltre per gestire gli errori della libreria deve essere gestita l'eccezione <i>ExcelException</i>

## Creazione di un foglio Excel
Per creare un Excel bisogna instanziare un oggetto <i>Writer</i> in questo modo:

```php
    $writer = new \SamagTech\ExcelLib\Writer($path);
```
oppure

```php
    $writer = (new \SamagTech\ExcelLib\Factory())->createWriter($path);
```

Il <i>Writer</i> ha bisogno del <i>path</i> dove verranno salvati i file ma accetta anche il nome del file e la lista delle colonne da ignorare la formattazione. Esempio di creazione

```php
    use SamagTech\ExcelLib\Writer;

    $writer = new Writer($path, $filename, $ignoreFieldsFormat);

    or

    $writer = new Writer($path, ignoreFieldsFormat: $ignoreFieldsFormat);

    or

    $writer = new Writer($path);
    $writer->setFilename($filename)->setIgnoreFieldsFormat($ignoreFieldsFormat);

```

è possibile cambiare anche il path a runtime

```php
    $writer->setPath($path);
```

### Utilizzo del Writer
Esempio di utilizzo base

```php
    $filePath = $writer->setHeader($headers)->setBody($body)->build();
```
La funzione <i>setHeader()</i> imposta l'intestazione dei fogli ed è facoltativa, la funzione <i>setBody()</i> setta i dati da inserire nel foglio e la funzione <i>build()</i> costruisce il foglio e restituisce il path con il nome del file.

### Definizione formattazione colonne
Utilizzando la funzione <i>setColumnDefinition(array $columnDefinition)</i> è possibile definire la formattazione delle colonne con un array dove la chiave è la chiave della colonna e il valore la tipologia di formattazione.

```php
    $columnDefinition = [
        'name'          => 'string',
        'age'           =>  'number',
        'perc_fat'      =>  'percentage'
    ];
```

<b>NB. Per ora sono gestiti sono number, string e percentage, di default vengono applicate solo le logiche di stringa e numero(i numeri vengono automaticamente formattati all'italiana X.XXX,XX e in rosso se negativi)</b>

### Costruzioni di Excel con fogli multipli
È possibile costruire file Excel con fogli multipli utilizzando sempre il metodo <i>build()</i> passandogli il parametro <i>true</i>, automaticamente verrà gestito il corpo come foglio multiplo.
Esempio di costruzione di foglio multiplo

```php
    $body = [
        // Foglio 1
        'sheet_1' => [
            // Riga 1
            [
                // Dati
            ],
            // Riga 2
            [
                // Dati
            ]
        ],
        // Foglio 2
        'sheet_2' => [
            // Riga 1
            [
                // Dati
            ],
            // Riga 2
            [
                // Dati
            ]
        ]
    ];

    $filePath = $writer->setBody($body)->build(true);
```

Riguardo all'intestazione (sempre opzionale)

```php

    $header = [
        // Intestazione foglio 1
        [
            // Stringhe che diverranno colonne
        ],
        // Intestazione foglio 2
        [
            // Stringhe che diverranno colonne
        ]
    ]

    or

    // In questo caso non essendo diviso l'intestazione verrà duplicata per ogni foglio
    $header = [
        // Stringhe che diverranno colonne
    ]
```

<br>
<br>
<br>

## Lettura di un foglio Excel
Per leggere un Excel bisogna instanziare un oggetto <i>Reader</i> in questo modo:

```php
    $reader = new \SamagTech\ExcelLib\Reader($path, $filename);
```
oppure

```php
    $reader = (new \SamagTech\ExcelLib\Factory())->createReader($path, $filename);
```

Il <i>Reader</i> ha bisogno del <i>path</i> e del <i>filename</i> per caricare il file. Il nome del file ed il path possono essere cambiati a runtime

```php
    $reader->setPath($path)

    or

    $reader->setFilename($filename)
```

### Utilizzo del Reader
Esempio di utilizzo base

```php
    $array = $reader->toArray();

    or

    $object = $reader->toObject();
```
È quindi possibile trasformare l'Excel in oggetto o array. Se il file che si sta caricando ha più di un foglio allora vengono caricati i fogli in questo modo

```php

    // Excel con foglio 1 e 2

    $array = $reader->toArray();

    // L'array sarà strutturato con il nome dei fogli dove gli spazi verranno modificati in underscore

    /**
     * $array = [
     *  'Foglio_1' => [
     *      [
     *          // Riga 1
     *      ],
     *      [
     *          // Riga 2
     *      ]
     *  ],
     *  'Foglio_2'   =>  [
     *      [
     *         // Riga 1
     *      ],
     *      [
     *          // Riga 2
     *      ]
     * ]
     *
     *
     * /

```

### Definzione custom delle chiavi in base alle colonne

È possibile dare una definizione custom per tradurre le colonne e con chiavi custom

```php

    $customColumnToKey = [
        'Nome'      => 'firstname',
        'Cognome'   => 'lastname',
        'Età'       =>  'age'
    ]

    $data = $reader->setColumnToKey($customColumnToKey)->toArray();

    /**
     * Es. Excel
     *
     * Nome         | Cognome   |   Età
     * Alessandro   | Marotta   |   25
     *
     * $data = [
     *      [
     *          'firstname' => 'Alessandro',
     *          'lastname' => 'Marotta',
     *          'age' => 25,
     *      ],
     *      // Altre righe
     * ]
     *
     * /

```
<b>NB. In questo caso vengono scartate le colonne non definite. È necessaria l'intestazione delle colonne nel foglio.</b>

### Lettura di fogli specifici

È possibile indicare implicitamente i fogli da caricare

```php

    $sheetNames = 'Foglio 1'; // or $sheetNames = ['Foglio 1', 'Foglio 2']

    $reader->setSheetNames($sheetNames)->toArray();
```

### Indicare che il foglio non ha intestazione
È possibile che i fogli non abbiano un intestazione quindi è opportuno indicarlo

```php
    $reader->setHasHeader(false)->toArray();
```

### Funzioni utili - Reader
<i>getNumSheet()</i> - Restituisce il numero di fogli effettivi caricati