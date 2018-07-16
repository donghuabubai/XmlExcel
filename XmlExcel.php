<?php

class XmlExcelObj
{
    public $worksheets = array();

    public function __construct($sWorksheetTitle)
    {
        $this->addsheet($sWorksheetTitle);
    }

    /**
     * @desc    添加分页
     * @param   $title
     */
    public function addsheet($title)
    {
        $this->worksheets[$title] = new WorkSheetObj($title);
    }

    /**
     * @desc    生成表格
     * @param   string $filename
     */
    public function generate($filename = 'excel-export')
    {
        header("Content-Type:application/force-download");
        header("Content-type:application/vnd.ms-csv");
        header("Content-Disposition:attachment; filename=" . $filename . ".xls");

        echo stripslashes("<?xml version=\"1.0\" encoding=\"UTF-8\"?\>\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">");
        foreach ($this->worksheets as $worksheet) {
            echo "\n<Worksheet ss:Name=\"" . $worksheet->sWorksheetTitle . "\">\n<Table>\n";
            $worksheet->printline();
            echo "</Table>\n</Worksheet>\n";
        }
        echo "</Workbook>";
    }

    /**
     * @desc    保存表格
     * @param   string $filename
     * @param   string $postfix
     * @param   $savePath
     */
    public function save($filename = 'excel-export', $postfix = 'xls', $savePath)
    {
        set_time_limit(0);
        //ini_set('memory_limit', '3072M');  //根据实际情况指定内存值

        $fileContent = stripslashes("<?xml version=\"1.0\" encoding=\"UTF-8\"?\>\n<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:html=\"http://www.w3.org/TR/REC-html40\">");
        foreach ($this->worksheets as $worksheet) {
            $fileContent .= "\n<Worksheet ss:Name=\"" . $worksheet->sWorksheetTitle . "\">\n<Table>\n";
            $fileContent .= $worksheet->getPrintLineStr();
            $fileContent .= "</Table>\n</Worksheet>\n";
        }
        $fileContent .= "</Workbook>";

        $myFile = fopen($savePath.$filename.'.'.$postfix, "w") or die("Unable to open file!");
        fwrite($myFile, $fileContent);
        fclose($myFile);
    }
}

class WorkSheetObj
{
    private $lines = array();
    public $sWorksheetTitle;

    public function __construct($sWorksheetTitle)
    {
        $this->setWorksheetTitle($sWorksheetTitle);
    }

	/**
     * @desc    设置单页表格名称
     * @param   $title
     */
    public function setWorksheetTitle($title)
    {
        $this->sWorksheetTitle = $title;
    }

	/**
     * @desc    往表格里添加数组数据
     * @param   $array
     */
    public function addRow($array)
    {
        foreach ($array as $v) {
            $cells = "";
            foreach ($v as $k => $v1) {
                $type = 'String';
                $v1 = htmlentities($v1, ENT_COMPAT, "UTF-8");
                $cells .= "<Cell><Data ss:Type=\"$type\">" . $v1 . "</Data></Cell>\n";

            }
            $this->lines[] = "<Row>\n" . $cells . "</Row>\n";
        }
    }

	/**
     * @desc    往表格里添加单行数据
     * @param   $array
     */
    public function addSingleRow($array)
    {
        $cells = "";
        foreach ($array as $index => $value) {
            $type = 'String';
            $value = htmlentities($value, ENT_COMPAT, "UTF-8");
            $cells .= "<Cell ss:Index=\"" . $index . "\" ><Data ss:Type=\"" . $type . "\">" . $value . "</Data></Cell>\n";
        }
        $this->lines[] = "<Row>\n" . $cells . "</Row>\n";
    }

	/**
     * @desc    往表格里添加带有合并单元格功能的单行数据
     * @param   $array
     */
    public function addSingleMergeRow($array)
    {
        $cells = "";
        foreach ($array as $index => $arr) {
            $type = 'String';
            $arr['value'] = htmlentities($arr['value'], ENT_COMPAT, "UTF-8");
            if (isset($arr['width']) && $arr['width'] > 0 && isset($arr['height']) && $arr['height'] > 0) {
                $cells .= "<Cell ss:MergeDown=\"" . $arr['height'] . "\" ss:MergeAcross=\"" . $arr['width'] . "\" ss:Index=\"" . $index . "\"><Data ss:Type=\"" . $type . "\">" . $arr['value'] . "</Data></Cell>\n";
            } elseif (isset($arr['width']) && $arr['width'] > 0) {
                $cells .= "<Cell ss:MergeAcross=\"" . $arr['width'] . "\" ss:Index=\"" . $index . "\"><Data ss:Type=\"" . $type . "\">" . $arr['value'] . "</Data></Cell>\n";
            } elseif (isset($arr['height']) && $arr['height'] > 0) {
                $cells .= "<Cell ss:MergeDown=\"" . $arr['height'] . "\" ss:Index=\"" . $index . "\"><Data ss:Type=\"" . $type . "\">" . $arr['value'] . "</Data></Cell>\n";
            } else {
                $cells .= "<Cell ss:Index=\"" . $index . "\"><Data ss:Type=\"" . $type . "\">" . $arr['value'] . "</Data></Cell>\n";
            }
        }
        $this->lines[] = "<Row>\n" . $cells . "</Row>\n";
    }

	/**
     * @desc    往表格里添加带有合并单元格功能的数组数据
     * @param   $array
     */
    public function addMultitermMergeRow($array)
    {
        foreach ($array as $value) {
            $this->addSingleMergeRow($value);
        }
    }

	/**
     * @desc    输出单行数据
     */
    public function printline()
    {
        foreach ($this->lines as $line) {
            echo $line;
        }
    }

	/**
     * @desc    单行数据字符拼接
     */
    public function getPrintLineStr()
    {
        $str = '';
        foreach ($this->lines as $line) {
            $str .= $line;
        }
        return $str;
    }

}

public function testFunc()
{
	$sheet_name = 'test';
	$XmlExcelObj = new XmlExcelObj($sheet_name);
	$exportData = [1,2,3,4,5];
	$XmlExcelObj->addsheet($sheet_name);
	$XmlExcelObj->worksheets[$sheet_name]->addRow(array($exportData));
	$XmlExcelObj->generate();
}

testFunc();


