# docx 模板工具類說明文檔
  
  **通過docx模板文本標記的形式進行文檔數據字符增刪改。**
  
---
 **log:**<br/>
  > begin-write by Chen Hai Quan (jacky) 2020/04/09  
  > end-write by Chen Hai Quan (jacky) 2020/04/09 

  > update-write by Chen Hai Quan (jacky) 2020/05/08
  >> 1.增加字符value值換行功能。
  > update-write by Chen Hai Quan (jacky) 2020/05/11
  >> 1.增加刪除整行功能。（包括整行节点）
  > update-write by Chen Hai Quan (jacky) 2020/05/12
  >> 1.增加刪除一段范围内节点功能。(局限一行內的範圍)
---

# 大綱

  <ul>
    <li>
     <a href="#basic">基本用法</a>
    </li>
    <li>
     <a href="#oneText">單個字符純文本替換功能</a>
    </li>
    <li>
     <a href="#arrayTblRow">數組遍歷生成tbl行功能</a>
    </li>
    <li>
     <a href="#deleteText">刪除一段文字功能</a>
    </li>
    <li>
      <a href="#deleteRow">刪除整行功能</a>
    </li>
    <li>
      <a href="#deleteNode">刪除一段节点功能</a>
    </li>
    <li>
     <a href="#strBr">字符串Value值換行標識</a>
    </li>
    
  </ul>

## <a id="basic">基本用法</a>
   
   ```javascript
    let doc = new hsDocxUtil({
       inputFileName:"test.docx",
       data:{
        name:"Chen Hai Quan",
        sex:"male",
        hobby:"play basketball",
        tbl:[
          {
             question:"who are you ?",
             answer:"I am super man."
          },
          {
             question:"are you super man ?",
             answer:"yes,I am."
          },
          {
             question:"can you go to hevaen ?",
             answer:"of course."
          }
        ],
        deleteBegin:"%deleteText%begin",
        deleteEnd:"%deleteText%end"
       }
    });

    //輸出buff
    let res=doc.getBuf();
    let buf;
    if(res.status){
      buf=res.data;
    }
    
    /**
     模板內容:
       我的名字是{name},性別是{sex},愛好為{hobby}.
       
       這是一段問答：
         {#tbl}
          問：{question}
          答：{answer}
         {#tbl}
       
       美國隊長{deleteBegin}是gay嗎？{deleteEnd}是hero嗎？
    */

    /**
     輸出文檔效果：
       我的名字是Chen Hai Quan,性別是male,愛好為play basketball.

       這是一段問答：
         問：who are you ?
         答：I am super man.
         問：are you super man ?
         答：yes,I am.
         問：can you go to hevaen ?
         答：of course.
      
       美國隊長是hero嗎？
    */
    //inputFileName-包含docx模板文當名的相對路徑-必需項<string>
   ```

## <a id="oneText">單個字符純文本替換功能</a>

   文本內容標記使用{key},key為傳入data對象的key.

   例：

    docx模板：

      我的名字是{name},性別是{sex},愛好為{hobby}
    
    data:
       {
         "name":"Chen Hai Quan",
         "sex":"male",
         "hobby":"play basketball..."
       };

## <a id="arrayTblRow">數組遍歷生成tbl行功能
    
   根據文檔表格的模板行生成多行的模板行內容。<br/>
   
  **模板:**
  
  | 問題      | 答案    | 得分  |
  | ----------|------- |-----  |
  | {#tbl}    |        |       |
  | {question}|{answer}|{score}|
  | {#tbl}    |        |       |

  ```javascript 
   data:
   {
       tbl:[
         {
             question:"who are you ?",
             answer:"I am super man.",
             score:100
          },
          {
             question:"are you super man ?",
             answer:"yes,I am.",
             score:100
          },
          {
             question:"can you go to hevaen ?",
             answer:"of course.",
             score:100
          }
       ]
   }  
  ```   

  **輸出效果：**

  |  問題                 |答案            |得分|
  |----                   |----           |----| 
  |who are you ?          |I am super man.|100 |
  |are you super man ?    |yes,I am.      |100 |
  |can you go to hevaen ? |of course.     |100 |
  
## <a id="deleteText">刪除一段文字功能(只清空文字)

  標記任意兩個
  <strong>key</strong>,
  <strong>value</strong>值為
  <strong>"%deleteText%begin"</strong>和
  <strong>"%deleteText%end"</strong><br/>
  **將要刪除的文字段包裹在此兩個key標記中**  
  
  ```javascript
    data:
    {
       key1:"%deleteText%begin",
       key2:"%deleteText%end" 
    }
    
    /**
      模板：
        美國隊長{key1}是gay嗎？{key2}是hero嗎？
    */
    
    /**
      效果：
        美國隊長是hero嗎？
    */

    //使用技巧1
    //如需通過數據判斷是否刪除此段文字,需刪除時如上傳值，不需刪除時如下傳空字符即可
    data:
     {
       key1:"",
       key2:""  
     } 
  ```
## <a id="deleteRow">删除一整行功能(包括元素)</a>
  標記任意兩個
  <strong>key</strong>,
  <strong>value</strong>值為
  <strong>"%deleteRow%begin"</strong>和
  <strong>"%deleteRow%end"</strong><br/>
  **將要刪除的文字段包裹在此兩個key標記中**  
  
  ```javascript
  data:
    {
       key1:"%deleteRow%begin",
       key2:"%deleteRow%end" 
    }
    
    /**
      模板：
        {key1}美國隊長是gay嗎？是hero嗎？{key2}
        答：是
    */
    
    /**
      效果：
        答：是
    */

    //使用技巧1
    //如需通過數據判斷是否刪除此行,需刪除時如上傳值，不需刪除時如下傳空字符即可
    data:
     {
       key1:"",
       key2:""  
     } 
  ```
## <a id="deleteNode">删除一段节点功能</a>
  標記任意兩個
  <strong>key</strong>,
  <strong>value</strong>值為
  <strong>"%deleteNode%begin"</strong>和
  <strong>"%deleteNode%end"</strong><br/>
  **將要刪除的文字段包裹在此兩個key標記中**  
  
  ```javascript
  data:
    {
       key1:"%deleteNode%begin",
       key2:"%deleteNode%end" 
    }
    
    /**
      模板：
        美國隊長是gay嗎？{key1}是hero嗎？{key2}
        答：是
    */
    
    /**
      效果：
        美國隊長是gay嗎？
        答：是
    */

    //使用技巧1
    //如需通過數據判斷是否刪除此行,需刪除時如上傳值，不需刪除時如下傳空字符即可
    data:
     {
       key1:"",
       key2:""  
     } 
  ```
## <a id="strBr">字符串Value值換行標識</a>

  如標記的數據value值內需通過加入標記實現字符串換行，在需換行的地方加上`<br>`即可實現換行



  

  


