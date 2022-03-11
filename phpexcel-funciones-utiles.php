<?php
//LIBRERIAS Y CONSTANTES
sc_include_library("sys", "PHPExcel", "PHPExcel.php", true, true);
define('WHITE','FFFFFF');



//EXCEL Valor 
function dev_excel_cell_val(&$objPHPExcel,$posicion,$valor,$color='',$textColor='',$bold=false){
    $posicion = count(explode(':',$posicion))>0 ? explode(':',$posicion)[0] : $posicion;
    $objPHPExcel->getActiveSheet()->setCellValue($posicion, trim(utf8_encode($valor)));

    if(isset($color{2})){
        dev_excel_cell_color($objPHPExcel,$posicion,$color);
    }

    if(isset($textColor{2}) || $bold){
        dev_excel_cell_text_color($objPHPExcel,$posicion,$textColor,$bold);
    }
}

//EXCEL Autosize
function dev_excel_col_auto_size(&$objPHPExcel){
    $sheet		  = $objPHPExcel->getActiveSheet();
    $cellIterator = $sheet->getRowIterator()->current()->getCellIterator();
    $cellIterator ->setIterateOnlyExistingCells( true );

    foreach( $cellIterator as $cell ) {
        $sheet->getColumnDimension( $cell->getColumn() )->setAutoSize( true );
        $sheet->getRowDimension   ($cell->getRow()     )->setRowHeight(-1);
    }
}

//EXCEL color (BG)
function dev_excel_cell_color(&$objPHPExcel,$posicion,$color='FFFFFF'){
    $objPHPExcel->getActiveSheet()->getStyle($posicion)->getFill()->applyFromArray(array(
        'type' => PHPExcel_Style_Fill::FILL_SOLID,
        'startcolor' => array(
            'rgb' => $color
        )
    ));
}

//EXCEL color (TEXTO)
function dev_excel_cell_text_color(&$objPHPExcel,$posicion,$textColor='FFFFFF',$bold=false){
    $styleArray = array(
        'font'  => array(
            'bold'  => $bold,
            'color' => array('rgb' => $textColor),
        ));

    $objPHPExcel->getActiveSheet()->getStyle($posicion)->applyFromArray($styleArray);
}

//EXCEL Titulo
function dev_excel_cell_titulo(&$objPHPExcel,$posicion,$valor,$color='',$textColor='',$fontSize=''){
    $sheet = $objPHPExcel->getActiveSheet();

    dev_excel_cell_val ($objPHPExcel,$posicion,$valor,$color,$textColor,true);

    if(count(explode(':',$posicion))>1){
        $objPHPExcel->setActiveSheetIndex(0)->mergeCells($posicion);
    }

    $style = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        )
    );

    $sheet->getStyle($posicion)->applyFromArray($style);

    if(isset($fontSize{1})){
        dev_excel_cell_font_size($objPHPExcel,$posicion,$fontSize);
    }

    if(isset($color{2})){
        dev_excel_cell_color($objPHPExcel,$posicion,$color);
    }


}

//EXCEL fontSize

function dev_excel_cell_font_size(&$objPHPExcel,$posicion,$fontSize){
    $objPHPExcel
        ->getActiveSheet()
        ->getStyle($posicion)
        ->getFont()
        ->setSize($fontSize);
}
