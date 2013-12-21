#include "stdafx.h"

#ifdef DLL_EXPORTS
#define DLL_API __declspec(dllexport) 
#else
#define DLL_API __declspec(dllimport) 
#endif

__declspec(dllexport) void Initialize();
__declspec(dllexport) void Uninitialize();

__declspec(dllexport) int PresentationCurrentSlideIndex();
__declspec(dllexport) int PresentationTotalSlidesCount();

__declspec(dllexport) void PresentationPreviousSlide();
__declspec(dllexport) void PresentationNextSlide();

__declspec(dllexport) int PresentationCurrentSlideName(unsigned char* dest, int destLen);
__declspec(dllexport) int PresentationCurrentSlideNote(unsigned char* dest, int destLen);

__declspec(dllexport) void RefreshPresentationSlidesThumbnail();
