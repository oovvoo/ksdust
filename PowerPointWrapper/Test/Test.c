// Test.cpp : Defines the entry point for the console application.
//

#include <stdio.h>
#include <stdlib.h>
#include <tchar.h>
#include "PowerPointWrapper.h"

int _tmain(int argc, _TCHAR* argv[])
{	
	Initialize();

	const int bufferLen = 1024;
	unsigned char* buffer = (unsigned char*)malloc(bufferLen * sizeof(unsigned char));
	buffer[0] = '\0';
	
	printf("presentation slide index: %d\n", PresentationCurrentSlideIndex());

	PresentationCurrentSlideNote(buffer, bufferLen);
	printf("presentation note: %s\n", buffer);

	PresentationCurrentSlideName(buffer, bufferLen);
	printf("slide name: %s\n", buffer);

	printf("saving presentation slides thumbnails...\n");
	RefreshPresentationSlidesThumbnail();

	/*printf("gonna goto previous slide...\n");
	system("pause");
	PresentationPreviousSlide();

	printf("gonna goto next slide...\n");
	system("pause");
	PresentationNextSlide();*/
	
	Uninitialize();

	printf("Test finished\n");
	system("pause");

	return 0;
}

