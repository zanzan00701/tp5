<?php
namespace app\index\controller;

use think\Loader;
use think\Controller;
use think\Db;
use think\PHPExcel_Reader_Excel2007;

class Index extends Controller
{
    public function index()
    {
        return "<a href='".url('excel')."'>导出</a>";
    }
    
	public function daochu()
	{
		 $path = dirname(__FILE__); //找到当前脚本所在路径
        Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
        $PHPExcel = new \PHPExcel(); //实例化
		  $xslist=db('iclass')->select(); // 查询学生内容
		  // echo"<pre>";
		  // print_r($xslist);exit;
		  foreach ($xslist as $key=> $v)
		  {
			$PHPExcel->createSheet();  //建sheet 值
            $PHPExcel->setactivesheetindex($key);    // 赋值
			$PHPSheet = $PHPExcel->getActiveSheet();
			$PHPSheet->setTitle($v['classname']); //给当前活动sheet设置名称
			$ar=array('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z');
			 $lists=Db::query('SHOW FULL COLUMNS from wx_users');
			  foreach($lists as $k=>$vo) {//获取当前表结构,赋给sheet的第一行
                    $comment=$vo['Comment']?$vo['Comment']:$vo['Field'];
                    $PHPSheet->setCellValue($ar[$k].'1',$comment);
                }
				 $users=db('users')->where("iclass='".$v['id']."'")->select(); //查询学生内容
				  $i=2;
                foreach($users as $k=>$vs){
                    $j=0;
                    foreach($vs as $k=>$v){
                        $PHPSheet->setCellValue($ar[$j].$i,$v);
                        $j++;
                    }
                        $i++;
                }
		$PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, "Excel2007"); //创建生成的格式
        header('Content-Disposition: attachment;filename="学生列表'.time().'.xlsx"'); //下载下来的表格名
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
	}
	}
	
    public function excel()
    {
        $path = dirname(__FILE__); //找到当前脚本所在路径
        Loader::import('PHPExcel.PHPExcel'); //手动引入PHPExcel.php
        Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 文件里面的PHPExcel_IOFactory这个类
        $PHPExcel = new \PHPExcel(); //实例化
        $iclasslist=db('iclass')->select();
        foreach($iclasslist as $key=> $v){
            $PHPExcel->createSheet();
            $PHPExcel->setactivesheetindex($key);
            $PHPSheet = $PHPExcel->getActiveSheet();
            $PHPSheet->setTitle($v['classname']); //给当前活动sheet设置名称
            $PHPSheet->setCellValue("A1", "编号")
                     ->setCellValue("B1", "姓名")
                     ->setCellValue("C1", "性别")
                     ->setCellValue("D1", "身份证号")
                     ->setCellValue("E1", "宿舍编号")
                     ->setCellValue("F1", "班级");//表格数据
            $userlist=db('users')->where("iclass=".$v['id'])->select();
            //echo db('users')->getLastSql();
            $i=2;
            foreach($userlist as $t)
            {
                $PHPSheet->setCellValue("A".$i, $t['id'])
                         ->setCellValue("B".$i, $t['username'])
                        ->setCellValue("C".$i, $t['sex'])
                        ->setCellValue("D".$i, $t['idcate'])
                        ->setCellValue("E".$i, $t['dorm_id'])
                        ->setCellValue("F".$i, $t['iclass']);
                        //表格数据
                $i++;
            }

        }
       // exit;
        $PHPWriter = \PHPExcel_IOFactory::createWriter($PHPExcel, "Excel2007"); //创建生成的格式
        header('Content-Disposition: attachment;filename="学生列表'.time().'.xlsx"'); //下载下来的表格名
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        $PHPWriter->save("php://output"); //表示在$path路径下面生成demo.xlsx文件
    }



    function importExecl($file='', $sheet=0){  
    
    Loader::import('PHPExcel.PHPExcel'); 
     Loader::import('PHPExcel.PHPExcel.IOFactory.PHPExcel_IOFactory'); //引入IOFactory.php 
    $objRead = new PHPExcel_Reader_Excel();   //建立reader对象  
    if(!$objRead->canRead($file)){  
        $objRead = new PHPExcel_Reader_Excel5();  
        if(!$objRead->canRead($file)){  
            die('No Excel!');  
        }  
    }  
  
    $cellName = array('A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z', 'AA', 'AB', 'AC', 'AD', 'AE', 'AF', 'AG', 'AH', 'AI', 'AJ', 'AK', 'AL', 'AM', 'AN', 'AO', 'AP', 'AQ', 'AR', 'AS', 'AT', 'AU', 'AV', 'AW', 'AX', 'AY', 'AZ');  
  
    $obj = $objRead->load($file);  //建立excel对象  
    $currSheet = $obj->getSheet($sheet);   //获取指定的sheet表  
    $columnH = $currSheet->getHighestColumn();   //取得最大的列号  
    $columnCnt = array_search($columnH, $cellName);  
    $rowCnt = $currSheet->getHighestRow();   //获取总行数  
  
    $data = array();  
    for($_row=1; $_row<=$rowCnt; $_row++){  //读取内容  
        for($column=0; $_column<=$columnCnt; $_column++){  
            $cellId = $cellName[$_column].$_row;  
            $cellValue = $currSheet->getCell($cellId)->getValue();  
             //$cellValue = $currSheet->getCell($cellId)->getCalculatedValue();  #获取公式计算的值  
            if($cellValue instanceof PHPExcel_RichText){   //富文本转换字符串  
                $cellValue = $cellValue->__toString();  
            }  
  
            $data[$_row][$cellName[$_column]] = $cellValue;  
        }  
    }  
  
    return $data;  
}  
}
