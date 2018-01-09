<?php
namespace app\index\controller;
use think\Controller;

class Test extends Controller{
	public function qq(){
		return  $this->fetch();
	}
	public function reg()
	{
		$xz=input('post.xz');
		// print_r($xz);exit;
			if($xz=='腾讯'){
				$email=input('post.email');
				$title=input('post.title');
				$content=input('post.content');
				$flag = sendMail($email,$title,$content);
			
			}else{
				$tomail=input('post.email');
				$title=input('post.title');
				$body=input('post.content');
				$flag = wymail($tomail,$title,$body);
				
			}
			if($flag){
		   $this->success('发送成功','Test/qq');
		}else{
		    $this->error('发送失败','Test/qq');
		}
		
	}
	
}

?>