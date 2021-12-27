<?php

namespace PhpDocxTemplate;

use DOMDocument;
use DOMElement;
use Exception;
use ZipArchive;
use RecursiveIteratorIterator;
use RecursiveDirectoryIterator;

/**
 * Class DocxDocument
 *
 * @package PhpDocxTemplate
 */
class DocxDocument
{
    private $path;
    private $tmpDir;
    private $document;
    private $zipClass;
    private $tempDocumentMainPart;
    private $tempDocumentRelations = [];

    /**
     * Construct an instance of Document
     *
     * @param string $path - path to the document
     *
     * @throws Exception
     */
    public function __construct(string $path)
    {
        if (file_exists($path)) {
            $this->path = $path;
            $this->tmpDir = sys_get_temp_dir() . "/" . uniqid("", true) . date("His");
            $this->zipClass = new ZipArchive();
            $this->extract();
        } else {
            throw new Exception("The template " . $path . " was not found!");
        }
    }

    /**
     * Extract (unzip) document contents
     */
    private function extract(): void
    {
        if (file_exists($this->tmpDir) && is_dir($this->tmpDir)) {
            $this->rrmdir($this->tmpDir);
        }

        mkdir($this->tmpDir);

        $this->zipClass->open($this->path);
        $this->zipClass->extractTo($this->tmpDir);
        $this->tempDocumentMainPart = $this->readPartWithRels($this->getMainPartName());
        $this->zipClass->close();

        $this->document = file_get_contents($this->tmpDir . "/word/document.xml");
    }

    /**
     * Get document main part
     *
     * @return string
     */
    public function getDocumentMainPart(): string
    {
        return $this->tempDocumentMainPart;
    }

    /**
     * Get the name of main part document (method from PhpOffice\PhpWord)
     *
     * @return string
     */
    private function getMainPartName(): string
    {
        $contentTypes = $this->zipClass->getFromName('[Content_Types].xml');

        $pattern = '~PartName="\/(word\/document.*?\.xml)" ' .
                   'ContentType="application\/vnd\.openxmlformats-officedocument' .
                   '\.wordprocessingml\.document\.main\+xml"~';

        $matches = [];
        preg_match($pattern, $contentTypes, $matches);

        return array_key_exists(1, $matches) ? $matches[1] : 'word/document.xml';
    }

    /**
     * Read document part (method from PhpOffice\PhpWord)
     *
     * @param string $fileName
     *
     * @return string
     */
    private function readPartWithRels(string $fileName): string
    {
        $relsFileName = $this->getRelationsName($fileName);
        $partRelations = $this->zipClass->getFromName($relsFileName);
        if ($partRelations !== false) {
            $this->tempDocumentRelations[$fileName] = $partRelations;
        }

        return $this->fixBrokenMacros($this->zipClass->getFromName($fileName));
    }

    /**
     * Get the name of the relations file for document part (method from PhpOffice\PhpWord)
     *
     * @param string $documentPartName
     *
     * @return string
     */
    private function getRelationsName(string $documentPartName): string
    {
        return 'word/_rels/' . pathinfo($documentPartName, PATHINFO_BASENAME) . '.rels';
    }


    /**
     * Finds parts of broken macros and sticks them together (method from PhpOffice\PhpWord)
     *
     * @param string $documentPart
     *
     * @return string
     */
    private function fixBrokenMacros(string $documentPart): string
    {
        return preg_replace_callback(
            '/\$(?:\{|[^{$]*\>\{)[^}$]*\}/U',
            function ($match) {
                return strip_tags($match[0]);
            },
            $documentPart
        );
    }

    /**
     * Get document.xml contents as DOMDocument
     *
     * @return DOMDocument
     */
    public function getDOMDocument(): DOMDocument
    {
        $dom = new DOMDocument();
        $dom->loadXML($this->document);
        return $dom;
    }

    /**
     * Update document.xml contents
     *
     * @param DOMDocument $dom - new contents
     */
    public function updateDOMDocument(DOMDocument $dom): void
    {
        $this->document = $dom->saveXml();
        file_put_contents($this->tmpDir . "/word/document.xml", $this->document);
    }

    /**
     * Fix table corruption
     *
     * @param string $xml - xml to fix
     *
     * @return DOMDocument
     */
    public function fixTables(string $xml): DOMDocument
    {
        $dom = new DOMDocument();
        $dom->loadXML($xml);
        $tables = $dom->getElementsByTagName('tbl');
        foreach ($tables as $table) {
            $columns = [];
            $columnsLen = 0;
            $toAdd = 0;
            $tableGrid = null;
            foreach ($table->childNodes as $el) {
                if ($el->nodeName == 'w:tblGrid') {
                    $tableGrid = $el;
                    foreach ($el->childNodes as $col) {
                        if ($col->nodeName == 'w:gridCol') {
                            $columns[] = $col;
                            $columnsLen += 1;
                        }
                    }
                } elseif ($el->nodeName == 'w:tr') {
                    $cellsLen = 0;
                    foreach ($el->childNodes as $col) {
                        if ($col->nodeName == 'w:tc') {
                            $cellsLen += 1;
                        }
                    }
                    if (($columnsLen + $toAdd) < $cellsLen) {
                        $toAdd = $cellsLen - $columnsLen;
                    }
                }
            }

            // add columns, if necessary
            if (!is_null($tableGrid) && $toAdd > 0) {
                $width = 0;
                foreach ($columns as $col) {
                    if (!is_null($col->getAttribute('w:w'))) {
                        $width += $col->getAttribute('w:w');
                    }
                }
                if ($width > 0) {
                    $oldAverage = $width / $columnsLen;
                    $newAverage = round($width / ($columnsLen + $toAdd));
                    foreach ($columns as $col) {
                        $col->setAttribute('w:w', round($col->getAttribute('w:w') * $newAverage / $oldAverage));
                    }
                    while ($toAdd > 0) {
                        $newCol = $dom->createElement("w:gridCol");
                        $newCol->setAttribute('w:w', $newAverage);
                        $tableGrid->appendChild($newCol);
                        $toAdd -= 1;
                    }
                }
            }

            // remove columns, if necessary
            $columns = [];
            foreach ($tableGrid->childNodes as $col) {
                if ($col->nodeName == 'w:gridCol') {
                    $columns[] = $col;
                }
            }
            $columnsLen = count($columns);

            $cellsLen = 0;
            $cellsLenMax = 0;
            foreach ($table->childNodes as $el) {
                if ($el->nodeName == 'w:tr') {
                    $cells = [];
                    foreach ($el->childNodes as $col) {
                        if ($col->nodeName == 'w:tc') {
                            $cells[] = $col;
                        }
                    }
                    $cellsLen = $this->getCellLen($cells);
                    $cellsLenMax = max($cellsLenMax, $cellsLen);
                }
            }
            $toRemove = $cellsLen - $cellsLenMax;
            if ($toRemove > 0) {
                $removedWidth = 0.0;
                for ($i = $columnsLen - 1; ($i + 1) >= $toRemove; $i -= 1) {
                    $extraCol = $columns[$i];
                    $removedWidth += $extraCol->getAttribute('w:w');
                    $tableGrid->removeChild($extraCol);
                }

                $columnsLeft = [];
                foreach ($tableGrid->childNodes as $col) {
                    if ($col->nodeName == 'w:gridCol') {
                        $columnsLeft[] = $col;
                    }
                }
                $extraSpace = 0;
                if (count($columnsLeft) > 0) {
                    $extraSpace = $removedWidth / count($columnsLeft);
                }
                foreach ($columnsLeft as $col) {
                    $col->setAttribute('w:w', round($col->getAttribute('w:w') + $extraSpace));
                }
            }
        }
        return $dom;
    }

    /**
     * Get total cells length
     *
     * @param array $cells - cells
     *
     * @return int
     */
    private function getCellLen(array $cells): int
    {
        $total = 0;
        foreach ($cells as $cell) {
            foreach ($cell->childNodes as $tc) {
                if ($tc->nodeName == 'w:tcPr') {
                    foreach ($tc->childNodes as $span) {
                        if ($span->nodeName == 'w:gridSpan') {
                            $total += intval($span->getAttribute('w:val'));
                            break;
                        }
                    }
                    break;
                }
            }
        }
        return $total + 1;
    }

    /**
     * Save the document to the target path
     *
     * @param string $path - target path
     */
    public function save(string $path): void
    {
        $rootPath = realpath($this->tmpDir);

        $zip = new ZipArchive();
        $zip->open($path, ZipArchive::CREATE | ZipArchive::OVERWRITE);

        $files = new RecursiveIteratorIterator(
            new RecursiveDirectoryIterator($rootPath),
            RecursiveIteratorIterator::LEAVES_ONLY
        );

        foreach ($files as $name => $file) {
            if (!$file->isDir()) {
                $filePath = $file->getRealPath();
                $relativePath = substr($filePath, strlen($rootPath) + 1);
                $zip->addFile($filePath, $relativePath);
            }
        }

        $zip->close();

        $this->rrmdir($this->tmpDir);
    }

    /**
     * Remove recursively directory
     *
     * @param string $dir - target directory
     */
    private function rrmdir(string $dir): void
    {
        $objects = scandir($dir);
        if (is_array($objects)) {
            foreach ($objects as $object) {
                if ($object != "." && $object != "..") {
                    if (filetype($dir . "/" . $object) == "dir") {
                        $this->rrmdir($dir . "/" . $object);
                    } else {
                        unlink($dir . "/" . $object);
                    }
                }
            }
            reset($objects);
            rmdir($dir);
        }
    }

    /**
     * Close document
     */
    public function close(): void
    {
        $this->rrmdir($this->tmpDir);
    }
}
