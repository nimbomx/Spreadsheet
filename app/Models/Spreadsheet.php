<?php

namespace App\Models;


class Spreadsheet
{
    public function __construct($file)
    {
        $this->fileName = $file;
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($this->fileName);
        $this->sheetCount = $spreadsheet->getSheetCount();
        $this->sheetNames = $spreadsheet->getSheetNames();
        $this->getSheets();
    }

    private function getSheets(){
        $this->sheets = [];
        foreach($this->sheetNames as $sheet){
            $this->sheets[] = ["Name" => $sheet];
        }
    }

    public function GetContentByName($name){
        $key = array_search($name, array_column($this->sheets, 'Name'));
        $this->GetContent($key);
    }
    public function GetTitlesByName($name){
        $key = array_search($name, array_column($this->sheets, 'Name'));
        $this->GetTitles($key);
    }
    public function GetTitles($key){
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($this->fileName);
        $worksheet = $spreadsheet->getSheet($key);
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn); 
        $titles = [];
        for ($col = 1; $col <= $highestColumnIndex; ++$col) {
            $value = $worksheet->getCellByColumnAndRow($col, 1)->getValue();
            $titles[] = $value;
        }
        $this->sheets[$key]['Titles'] = $titles;
    }
    public function GetContent($key){
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($this->fileName);
        $worksheet = $spreadsheet->getSheet($key);
        $this->sheets[$key]['Data'] = []; 
        $title=true;
        foreach ($worksheet->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(FALSE);
            $cells = [];
            foreach ($cellIterator as $cell) {
                $cells[] = $cell->getValue();
            }
            if($title){
                $this->sheets[$key]['Titles'] = $cells;
                $title=false;
            }else{
                $this->sheets[$key]['Data'][] = $cells;
            }
        }
        //dd($worksheet->getHighestRow());
        //dd($worksheet->getHighestColumn());
    }
    public function Sheets(){
        return $this->sheets;
    }

    public function json(){
        return json_encode($this->sheets);
    }
}
