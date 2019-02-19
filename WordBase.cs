using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;

//Chr(10) 换行 Chr(13) 回车符
public class MyWordBase
{
    Application wordApp = new Application(); //Word应用 Application对象 Document wordDocument; //Word文档 Document对象
    Document wordDocument = null;
    int startPosition; //当前Word开始位置索引数组列表
    int endPosition;  //当前Word结尾位置索引数组列表
    Range rng = null; //Word范围 Range对象  
    Bookmark bookMark = null;
    Int16 wordstatus = 0; //word文档状态，0为已关闭，1为打开着


    public Application WordApp
    {
        get
        {
            return wordApp;
        }
    }
    public Document WordDocument
    {
        set
        {
            wordDocument = value;
        }
        get
        {
            return wordDocument;
        }
    }
    public int StartPosition
    {
        get
        {
            return startPosition;
        }
    }
    public int EndPosition
    {
        get
        {
            return endPosition;
        }
    }
    public Range Rng
    {
        get
        {
            return rng;
        }
    }
    public Int16 Status
    {
        set
        {
            wordstatus = value;
        }
        get
        {
            return wordstatus;
        }
    }
    #region 打开或新建文档
    public MyWordBase()
    {
        wordDocument = wordApp.Documents.Add();  //新建空文档
        Status = 1;
    }
    public MyWordBase(string FullFilePath, EnumCollections.OpenWay OpenWay)
    {
        FullFilePath = TrimUpper(FullFilePath);
        if (Convert.ToInt16(OpenWay) == 0)
        {
            wordDocument = wordApp.Documents.Add(FullFilePath);  //已现有文档为模板新建文档
        }
        else if (Convert.ToInt16(OpenWay) == 1)
        {
            wordDocument = wordApp.Documents.Open(FullFilePath);  //打开现有文档
        }
        else if (Convert.ToInt16(OpenWay) == 2)
        {
            wordDocument = wordApp.Documents.Open(FullFilePath, ReadOnly: true); //已只读方式打开现有文档
        }
        Status = 1;
    }
    #endregion

    #region 关闭文档
    public void CloseWordDocument(EnumCollections.Save Save)
    {
        if (wordstatus == 1)
        {
            if (Convert.ToInt16(Save) == 0)
            {
                wordDocument.Close();  //不保存直接关闭文档
            }
            else if (Convert.ToInt16(Save) == 1)
            {
                wordDocument.Close(SaveChanges: true); //默认保存关闭文档
            }
        }
        wordstatus = 0;
    }
    #endregion

    #region 保存文档
    public void SaveWordDocument(string FullFilePath = "")
    {
        FullFilePath = TrimUpper(FullFilePath);
        if (FullFilePath != "" && wordDocument.Name == "")
        {
            wordApp.Documents[FullFilePath].Save(); //文档未保存过，设置了路径，保存      
        }
        else if (FullFilePath != "" && wordDocument.Name != "")
        {
            wordDocument.SaveAs(FullFilePath); //文档保存过，设置了路径，另存
        }
        else if (wordDocument.Name != "")
        {
            wordDocument.Save();  //文档保存过，没有设置路径，保存
        }
    }
    public void Save()
    {
        wordDocument.Save();
    }
    #endregion

    #region 关闭Word组件对象
    public void CloseWordApp()
    {
        if (wordstatus == 1)
        {
            CloseWordDocument(0);
            wordstatus = 0;
        }
        wordApp.Quit();
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordDocument);
        System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
    }
    #endregion

    #region 查找文本位置并获取查找对象的开始，结尾位置索引,存入数组 暂时只能搜索并保存数组列表单次，第二次搜到直接跳出方法
    public void FindPosition(string findtext)
    {
        startPosition = 0;
        endPosition = 0;
        findtext = TrimUpper(findtext);
        foreach (Paragraph paragraph in wordDocument.Paragraphs)  //循环文档页，查找字符串
        {
            rng = paragraph.Range;
            if (rng.Find.Execute(findtext))
            {

                startPosition = Rng.Start;

                endPosition = Rng.End;
            }
        }
        rng = null;
    }
    #endregion


    #region 返回范围内文本
    public string ReturnText(int startPosition, int endPosition)
    {
        string returntext = "";
        rng = WordDocument.Range(startPosition, endPosition);
        returntext = rng.Text.Trim();
        return returntext;
    }
    #endregion

    #region 在范围中插入文本
    public int InsertText(int startPosition, int endPosition, string insertText)
    {
        int Ok = 0; //判断是否插入成功
        if (endPosition - startPosition == 0)
        {
            rng = wordDocument.Range(startPosition, endPosition);
            rng.Text = insertText;
            Ok = 1;     //单点插入，返回1
        }
        else if (insertText.Length <= (endPosition - startPosition))
        {
            rng = wordDocument.Range(startPosition, endPosition);
            rng.Text = insertText;
            Ok = 1; //范围插入，返回1
        }
        else if (insertText.Length > (endPosition - startPosition))
        {
            Ok = 0; //插入字符串长度大于范围，返回0，以免破坏文档结构
        }
        return Ok;
    }
    #endregion

    #region 获取书签内文本
    public string GetBookMarkText(string bookmarkname)
    {
        bookMark = wordDocument.Bookmarks[bookmarkname];
        return bookMark.Range.Text.Trim();
    }
    #endregion

    #region 替换书签内文本
    public void ReplaceBookMarkText(string bookmarkname, string replacetext)
    {
        bookMark = wordDocument.Bookmarks[bookmarkname];
        bookMark.Range.Text = replacetext;

    }
    #endregion

    #region 设置范围书签，以开始、结束范围设置
    public int SetBookMarkText(int startposition, int endposition, string bookmarkname)
    {
        int o = 0;
        Range sRange = WordDocument.Range(startposition, endposition);
        try
        {
            WordDocument.Bookmarks.Add(bookmarkname, sRange);
            o = 1;

        }
        catch
        {
            o = 0;
        }
        return o;
    }
    #endregion

    #region 将传入字符串大写并去除两侧空格
    private string TrimUpper(string InputStr)
    {
        return InputStr.Trim().ToUpper();
    }
    #endregion

    #region 强制关闭Word进程
    public int KillWord() //调用方法，传参
    {
        int o = 0;
        try
        {
            Process[] thisproc = Process.GetProcessesByName("WINWORD");
            //thisproc.lendth:名字为进程总数
            if (thisproc.Length > 0)

            {
                for (int i = 0; i < thisproc.Length; i++)
                {
                    if (!thisproc[i].CloseMainWindow()) //尝试关闭进程 释放资源
                    {
                        thisproc[i].Kill(); //强制关闭
                        o = 1;

                    }
                }
            }
            else
            {
                o = 0;
            }
        }
        catch //出现异常，表明 kill 进程失败
        {
            o = -1;
        }
        return o;
        #endregion
    }
}
