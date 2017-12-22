<?php
    header("content-type:text/html;charset=utf-8");
    $Id=$_POST['Id'];
    $Name=$_POST['Name'];
    $Sex=$_POST['Sex'];
    $Tel=$_POST['Tel'];
    $Native=$_POST['Native'];
    $ElabGroup=$_POST['ElabGroup'];
    $Pro=$_POST['Pro'];
    $DlutClass=$_POST['DlutClass'];
    $StuPosition=$_POST['StuPosition'];
    $Community=$_POST['Community'];
    $FreeTime=$_POST['FreeTime'];
    $Email=$_POST['Email'];
    $Experience=$_POST['Experience'];
    $TimeWeek=$_POST['TimeWeek'];
    $Expect=$_POST['Expect'];
    $SelfEvaluation=$_POST['SelfEvaluation'];

    if($Id == ""||$Name == ""||$Sex == ""||$Tel == ""||$Native == ""||$ElabGroup == ""||$Pro == ""||$DlutClass == ""||$StuPosition == ""||$Community == ""||$FreeTime == ""||$Email == ""||$Experience == ""||$Expect == ""||$TimeWeek == ""||$SelfEvaluation == ""){
        echo "<script>alert('信息输入不完全');</script>";
        echo "<script>history.go(-1);</script>";
    }
    else if(strlen($Id) != 9)
    {
        echo "<script>alert('请输入正确学号');</script>";
        echo "<script>history.go(-1);</script>";
    }
    else if(strlen($Tel) != 11)
    {
        echo "<script>alert('请输入正确手机号码');</script>";
        echo "<script>history.go(-1);</script>";
    }
    else
    {
        //通过php连接到mysql数据库
        $conn=new mysqli("localhost","admin","admin123456","NewPartner");

        //通过php进行insert操作
        $conn -> set_charset('utf8');

        $sqlinsert="replace into newperson(Id,Name,Sex,Tel,Native,ElabGroup,Professor,DlutClass,StuPosition,Community,FreeTime,Email,Experience,TimeWeek,Expect,SelfEvaluation) values('{$Id}','{$Name}','{$Sex}','{$Tel}','{$Native}','{$ElabGroup}','{$Pro}','{$DlutClass}','{$StuPosition}','{$Community}','{$FreeTime}','{$Email}','{$Experience}','{$TimeWeek}','{$Expect}','{$SelfEvaluation}')";

        if ($conn->query($sqlinsert) == TRUE) {
            echo "<script>alert('{$Id}' + '_{$Name}' + '_{$Sex}' + '_{$ElabGroup}' + ' 报名成功！');</script>";
            echo "<script>history.go(-1);</script>";
        }
        else {
            echo "Error: " . $sqlinsert . "<br>" . $conn->error;
        }

        //释放连接资源
        $conn->close();
    }

?>
