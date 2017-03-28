<?php

// namespace app\common\adapter;

// use \think\Db;

class WordAdapter
{
    public function exportWord($status, $wordName='数据库文档')
    {
        date_default_timezone_set('PRC');
        include "vendor/autoload.php";
        $phpWord = new \PhpOffice\PhpWord\PhpWord();

        $section = $phpWord->addSection();

        $phpWord->setDefaultFontName('宋体');
        $phpWord->addTitleStyle(1, array('size' => 16, 'bold' => true, 'lineHeight' => 2, 'name' => 'Arial'));
        $phpWord->addTitleStyle(2, array('size' => 15, 'bold' => true));
        $phpWord->addTitleStyle(3, array('size' => 12, 'bold' => true));
        $phpWord->addTitleStyle(4, array('size' => 11, 'bold' => true));
        $phpWord->addTitleStyle(5, array('size' => 10, 'bold' => true));

        $section->addText('数据库表结构说明文档', ['size' => 20, 'bold' => true], ['align' => 'center']);
        $section->addTextBreak(2);
        $section->addText('目         录', ['size' => 10.5], ['align' => 'center']);
        $toc = $section->addTOC(['size' => 10.5, 'name' => 'Times New Roman']);
        $section->addTextBreak(1);

        // $status = Db::query('SHOW table status');
        $name = array_column($status, 'Name');
        $comment = array_column($status, 'Comment');

        foreach ($name as $k => $v) {
            // $section->addText($v . ' ' . $comment[$k], 1);
            $section->addTitle($v, 1);
            $section->addText('描述：' . $comment[$k], ['size' => 10.5]);
            $table = $section->addTable(array('borderSize' => 6, 'cellMargin' => 100, 'valign' => 'center'));
            // $table->setStyle(array('border' => 10));
            $table->addRow();
            $table->addCell(710)->addText('序号', ['bold' => true], array('align' => 'center'));
            $table->addCell(1534)->addText('字段名称', ['bold' => true], array('align' => 'center'));
            $table->addCell(2730)->addText('字段描述', ['bold' => true], array('align' => 'center'));
            $table->addCell(1534)->addText('字段类型', ['bold' => true], array('align' => 'center'));
            $table->addCell(875)->addText('长度', ['bold' => true], array('align' => 'center'));
            $table->addCell(875)->addText('允许空', ['bold' => true], array('align' => 'center'));
            $table->addCell(1301)->addText('缺省值', ['bold' => true], array('align' => 'center'));

            foreach ($status[$k]['fields'] as $key => $value) {
                $table->addRow();
                $table->addCell(710)->addText($key+1, [], array('align' => 'center'));
                $table->addCell(1534)->addText($value['Field']);
                $value['Type'] = str_replace('(', '|', $value['Type']);
                $value['Type'] = str_replace(')', '|', $value['Type']);
                $type = explode('|', $value['Type']);

                count($type) > 1 && $long = $type[1];
                $table->addCell(2730)->addText($value['Comment']);
                $table->addCell(1534)->addText($type[0]);
                $table->addCell(875)->addText($long, [], array('align' => 'right'));
                $table->addCell(875)->addText($value['Null'] == 'NO' ? '' : '√', [], array('align' => 'center'));
                $table->addCell(1301)->addText($value['Default']);
            }
            $section->addTextBreak(2);
            // var_dump($comment[$k]);die;
        }
        // die;

        // $table = $section->addTable();
        // for ($i = 1; $i <= 8; $i++) {
        //     $table->addRow();
        //     $table->addCell(2000)->addText('');
        //     $table->addCell(2000)->addText('');
        //     $table->addCell(2000)->addText('');
        //     $table->addCell(2000)->addText('');
        //     $text = (0== $i % 2) ? 'X' : '';
        //     $table->addCell(500)->addText($text);
        // }

        $objWriter = \PhpOffice\PhpWord\IOFactory::createWriter($phpWord, 'Word2007');
        $objWriter->save("$wordName.docx");
        return $objWriter;
    }
}
