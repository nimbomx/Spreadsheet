<?php

namespace App\Http\Controllers;

use App\Models\Spreadsheet as ModelsSpreadsheet;
use Illuminate\Http\Request;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class HelloWorldController extends Controller
{
    public function write(){
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Hello World !');
        $writer = new Xlsx($spreadsheet);
        $writer->save('hello world.xlsx');
    }

    public function read(){
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load("hello world.xlsx");
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A2', 'Second Line');
        $writer = new Xlsx($spreadsheet);
        $writer->save('Secondline.xlsx');
    }

    public function show(){
        $S = new ModelsSpreadsheet("Secondline.xlsx");
        $S->GetContentByName("ASIGNACION (Shipping Window)");
        $S->GetTitlesByName("CUADROS");
        //$S->GetContent(0);
        //$S->GetTitles(1);
        dd($S->Sheets());
        return $S->Sheets();
    }
}
