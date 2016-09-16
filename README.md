#Instructions#
* 按下空格以使用补全，但只在你至少键入一个字母时有效

* HighLights存储着所有关键字和它们的颜色，格式如下<br>
``` VisualBasic
   ->  
   R G B  
   关键字1  
   关键字2  
   关键字3  
   ...  
```

* Statements存储着所有的关键字的从属关系，格式如下<br>
``` VisualBasic
    关键字1  
    ->  
    属于关键字1的子关键字1  
    ->  
    属于上个子关键字的关键字  
    <-  
    属于关键字1的子关键字2  
    <-  
    关键字2  
    ...  
```
```
      test.test1.test2
      test.test1.test3
      test.test4.test5 
```
      
在文件中的表述形式为

```
      test
      ->
      test1
      ->
      test2
      test3
      <-
      test4
      ->
      test5
      <-
      <-
```

希望能帮忙完善Statements，因为我也不清楚究竟有多少指令QwQ  

暂时测试是没有问题的了，但是长远的测试还没有怎么做  
菜单简陋到能让人吐槽死= =  
还有，我懒得写改字体的对话框，我觉得微软雅黑挺好看的  

>Powered by Prunoideae  
 2016.9.16
