// PowerPointWrapper.cpp : Defines the exported functions for the DLL application.
//

#include "stdafx.h"
#include <stdio.h>
#include <string.h>
#include "disphelper.h"
#include "PowerPointWrapper.h"

#define HR_TRY(func) if (FAILED(func)) {/* printf("\n## Fatal error on line %d.\n", __LINE__);*/ goto cleanup; }
#define PathBufferLen 1024

IDispatch* pptApp = NULL;

int encode(const wchar_t* wstr, unsigned int codePage, unsigned char* dest, int destLen);
void getProgramDirectory(wchar_t* dest, int destLength);

/*void debug(unsigned char* dest, int destLen) {
	FILE* f = _wfopen(L"a.txt", L"w");
	fprintf(f, "dest = %x, destLen = %d, strlen = %d\n", dest, destLen, strlen(dest));
}*/

// 初始化 COM 函式庫，在進行任何操作前必須先呼叫此方法。
void Initialize()
{
	dhInitialize(TRUE);
		
	// this is useful when debugging. When switched on,
	// any error will shown in a message box
	//dhToggleExceptions(TRUE);

	dhToggleExceptions(FALSE);
	dhCreateObject(L"PowerPoint.Application", NULL, &pptApp);
}

// 釋放 COM 函式庫的資源，不再使用這個 package 的功能時可呼叫此方法釋放資源。
void Uninitialize() {
	SAFE_RELEASE(pptApp);
	dhToggleExceptions(FALSE);
	dhUninitialize(TRUE);
}

// 判斷是否有開啟的簡報檔案
int HasActivePresentation() {
	IDispatch* activePresentation = NULL;
	HRESULT hResult = dhGetValue(L"%o", &activePresentation, pptApp, L".ActivePresentation");

	return activePresentation != NULL;
}

// 判斷是否正在放映投影片
int HasSlideShowWindow() {
	IDispatch* slideShowWindow = NULL;
	HRESULT hResult = dhGetValue(L"%o", &slideShowWindow, pptApp, L".ActivePresentation.SlideShowWindow");

	return slideShowWindow != NULL;
}

// 傳回現在正在放映的投影片 index，若發生任何錯誤，傳回 -1
int PresentationCurrentSlideIndex()
{
	if (!HasActivePresentation() || !HasSlideShowWindow())
	{
		return -1;
	}

	int currentSlideInndex;
	HRESULT hResult = dhGetValue(L"%d", &currentSlideInndex, pptApp, L".ActivePresentation.SlideShowWindow.View.Slide.SlideIndex");

	return currentSlideInndex;
}

// 傳回現在正在放映的簡報所含的投影片數量，若發生任何錯誤，傳回 -1
int PresentationTotalSlidesCount()
{
	if (!HasActivePresentation())
	{
		return -1;
	}

	int totalSlides;
	HRESULT hResult = dhGetValue(L"%d", &totalSlides, pptApp, L".ActivePresentation.Slides.Count");

	return totalSlides;
}

// 控制簡報：前往上一張投影片
void PresentationPreviousSlide()
{
	if (!HasActivePresentation() || !HasSlideShowWindow() || PresentationCurrentSlideIndex() <= 1)
	{
		return;
	}

	dhCallMethod(pptApp, L".ActivePresentation.SlideShowWindow.View.Previous");
}

// 控制簡報：前往下一張投影片
void PresentationNextSlide()
{
	
	if (!HasActivePresentation() || !HasSlideShowWindow() 
		|| PresentationCurrentSlideIndex() < 1 || PresentationCurrentSlideIndex() >= PresentationTotalSlidesCount())
	{
		return;
	}

	dhCallMethod(pptApp, L".ActivePresentation.SlideShowWindow.View.Next");
}

// 傳回目前放映的投影片的名稱，return string length stored in dest
int PresentationCurrentSlideName(unsigned char* dest, int destLen)
{
	if (!HasActivePresentation() || !HasSlideShowWindow() || PresentationCurrentSlideIndex() < 1)
	{
		return -1;
	}

	LPTSTR szText = NULL;

	if (dest == NULL || destLen < 0)
	{
		return 0;
	}

	HRESULT hResult = dhGetValue(L"%T", &szText, pptApp, L".ActivePresentation.SlideShowWindow.Presentation.Name");

	memset(dest, 0x0, destLen);
	int sizeUsed = encode(szText, CP_UTF8, dest, destLen);
	if (szText != NULL) {
		dhFreeString(szText);
	}

	if (sizeUsed <= 0) {
		return 0;
	}

	return sizeUsed - 1;
}

// 傳回目前放映的投影片的備忘稿，return string length stored in dest
int PresentationCurrentSlideNote(unsigned char* dest, int destLen)
{
	if (!HasActivePresentation() || !HasSlideShowWindow() || PresentationCurrentSlideIndex() < 1)
	{
		return 0;
	}

	LPTSTR szText = NULL;

	if (dest == NULL || destLen < 0)
	{
		return 0;
	}

	//int targetSlideIndex = 1;
	//HR_TRY(dhGetValue(L"%T", &szText, pptApp, L".ActivePresentation.Slides(%d).NotesPage.Shapes.Placeholders(%d).TextFrame.TextRange.Text", targetSlideIndex, 2));
	HRESULT hResult = dhGetValue(L"%T", &szText, pptApp, L".ActivePresentation.SlideShowWindow.View.Slide.NotesPage.Shapes.Placeholders(%d).TextFrame.TextRange.Text", 2);

	memset(dest, 0x0, destLen);
	int sizeUsed = encode(szText, CP_UTF8, dest, destLen);
	if (szText != NULL) {
		dhFreeString(szText);
	}

	//debug(dest, destLen);

	if (sizeUsed <= 0) {
		return 0;
	}

	return sizeUsed - 1;
}

// deletes all file in a foler, and create thumbnails of active slideshow.
// note: failed execution of this function will NOT recover any files previous existed in the folder.
void RefreshPresentationSlidesThumbnail()
{
	wchar_t curPath[PathBufferLen];
	getProgramDirectory(curPath, PathBufferLen);

	const LPCTSTR DirectoryName = L"presthumb\\";
	wcsncat(curPath, DirectoryName, PathBufferLen - wcslen(DirectoryName));

	const TargetThumbnailWidth = 1024;
	int thumbnailWidth, thumbnailHeight;

	HRESULT hResult = dhGetValue(L"%d", &thumbnailWidth, pptApp, L".ActivePresentation.PageSetup.SlideWidth");
	hResult = dhGetValue(L"%d", &thumbnailHeight, pptApp, L".ActivePresentation.PageSetup.SlideHeight");

	double aspectRatio = thumbnailWidth / (double)thumbnailHeight;
	thumbnailWidth = TargetThumbnailWidth;
	thumbnailHeight = (int)(TargetThumbnailWidth / aspectRatio);

	// deletion may not be successful if some files in the directory is still being used?
	// this delete command will not delete any directory resides in
	wchar_t command[PathBufferLen];
	command[0] = L'\0';
	wcsncat(command, L"del ", PathBufferLen - 4);
	wcsncat(command, curPath, PathBufferLen - wcslen(curPath));
	wcsncat(command, L"*.* / s / q", PathBufferLen - 11);
	_wsystem(command);

	CreateDirectory(curPath, NULL);

	int slidesCount = PresentationTotalSlidesCount();
	if (slidesCount < 1)
	{
		return;
	}

	wchar_t itowBuffer[10];
	wchar_t* curPathEndPtr = curPath + wcslen(curPath);

	for (int i = 1; i <= slidesCount; i++)
	{
		*curPathEndPtr = L'\0';
		_itow_s(i, itowBuffer, 8, 10);

		//wcsncat(curPathEndPtr, L"s", MAX_PATH - 1);
		wcsncat(curPathEndPtr, itowBuffer, PathBufferLen - wcslen(itowBuffer));
		wcsncat(curPathEndPtr, L".png", PathBufferLen - 4);
		wprintf(L"%s\n", curPath);
		dhCallMethod(pptApp, L".ActivePresentation.Slides(%d).Export(%S, %S, %d, %d)", i, curPath, L"png", thumbnailWidth, thumbnailHeight);
	}
	
}

void getProgramDirectory(wchar_t* dest, int destLength)
{
	int pathLength = GetModuleFileName(NULL, dest, destLength);
	wchar_t *destPtr = dest + pathLength / 2;

	while (wcschr(destPtr, L'\\'))
	{
		destPtr = wcschr(destPtr, L'\\');
		destPtr++;
	}

	*(destPtr) = L'\0';
}

// return size used
int encode(const wchar_t* wstr, unsigned int codePage, unsigned char* dest, int destLen)
{
	if (dest == NULL || destLen < 0)
	{
		return 0;
	}

	int sizeNeeded = WideCharToMultiByte(codePage, 0, wstr, -1, NULL, 0, NULL, NULL);
	if (destLen < sizeNeeded)
	{
		return 0;
	}

	WideCharToMultiByte(codePage, 0, wstr, -1, dest, sizeNeeded, NULL, NULL);

	return sizeNeeded;
}

/*
wchar_t* decode(const char* encodedStr, unsigned int codePage)
{
	int sizeNeeded = MultiByteToWideChar(codePage, 0, encodedStr, -1, NULL, 0);
	wchar_t* decodedStr = (char *)malloc(sizeNeeded * sizeof(wchar_t));
	MultiByteToWideChar(codePage, 0, encodedStr, -1, decodedStr, sizeNeeded);
	return decodedStr;
}*/
