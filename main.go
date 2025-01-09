package main

import (
	"fmt"
	"log"

	"github.com/go-ole/go-ole"
	"github.com/go-ole/go-ole/oleutil"
)

func main() {
	defer func() {
		if r := recover(); r != nil {
			fmt.Println("捕获到panic:", r)
		}
	}()

	ole.CoInitialize(0)
	defer ole.CoUninitialize()

	unknown, err := oleutil.CreateObject("KWPS.Application") // MSWord:"Word.Application"
	if err != nil {
		log.Fatal(err)
	}
	wordApp, err := unknown.QueryInterface(ole.IID_IDispatch)
	if err != nil {
		log.Fatal(err)
	}
	defer wordApp.Release()

	oleutil.PutProperty(wordApp, "Visible", false)

	documents := oleutil.MustGetProperty(wordApp, "Documents").ToIDispatch()
	doc := oleutil.MustCallMethod(documents, "Open", "C:\\Users\\Administrator\\Desktop\\go_ole\\test.docx").ToIDispatch()
	defer doc.Release()

	bookmarks := oleutil.MustGetProperty(doc, "Bookmarks").ToIDispatch()
	defer bookmarks.Release()

	count := oleutil.MustGetProperty(bookmarks, "Count").Val
	if count == 0 {
		fmt.Println("There is no bookmark in this document")
		return
	}

	for i := 1; i <= int(count); i++ {
		bookmark := oleutil.MustGetProperty(bookmarks, "Item", i).ToIDispatch()
		defer bookmark.Release()

		name := oleutil.MustGetProperty(bookmark, "Name").ToString()

		// 获取书签内容
		wRange := oleutil.MustGetProperty(bookmark, "Range").ToIDispatch()
		defer wRange.Release()

		text := oleutil.MustGetProperty(wRange, "Text").ToString()
		fmt.Println(i, ":", name, ":", text)
	}

	// 关闭文档
	oleutil.MustCallMethod(doc, "Close", false)
	// 退出Word应用程序
	oleutil.MustCallMethod(wordApp, "Quit")
}
