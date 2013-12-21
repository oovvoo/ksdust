package main

import (
	"bufio"
	"code.google.com/p/go.net/websocket"
	"compress/gzip"
	"fmt"
	"io"
	"io/ioutil"
	"ksdust/pptctrl"
	"net/http"
	"os"
	"strings"
)

var (
	// 要求的縮圖不存在時傳回的 png 圖片 (1x1, 黑色)
	_ThumbnailNotFoundImage = []byte{0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0x00, 0x00, 0x00, 0x0d, 0x49, 0x48, 0x44, 0x52, 0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, 0x08, 0x02, 0x00, 0x00, 0x00, 0x90, 0x77, 0x53, 0xde, 0x00, 0x00, 0x00, 0x01, 0x73, 0x52, 0x47, 0x42, 0x00, 0xae, 0xce, 0x1c, 0xe9, 0x00, 0x00, 0x00, 0x04, 0x67, 0x41, 0x4d, 0x41, 0x00, 0x00, 0xb1, 0x8f, 0x0b, 0xfc, 0x61, 0x05, 0x00, 0x00, 0x00, 0x09, 0x70, 0x48, 0x59, 0x73, 0x00, 0x00, 0x12, 0x74, 0x00, 0x00, 0x12, 0x74, 0x01, 0xde, 0x66, 0x1f, 0x78, 0x00, 0x00, 0x00, 0x0c, 0x49, 0x44, 0x41, 0x54, 0x18, 0x57, 0x63, 0x60, 0x60, 0x60, 0x00, 0x00, 0x00, 0x04, 0x00, 0x01, 0x5c, 0xcd, 0xff, 0x69, 0x00, 0x00, 0x00, 0x00, 0x49, 0x45, 0x4e, 0x44, 0xae, 0x42, 0x60, 0x82}
)

func main() {
	go pptctrl.StartEventLoop()
	fmt.Println("starting...")

	http.Handle("/remote", websocket.Handler(remoteControl))
	http.HandleFunc("/", clientInterface)
	http.HandleFunc("/presthumbnail/", slideThumbnail)

	fmt.Println("在網頁的任何地方按下滑鼠左鍵或是觸摸螢幕")

	err := http.ListenAndServe(":6611", nil)
	if err != nil {
		panic("ListenAndServe: " + err.Error())
	}
}

func clientInterface(w http.ResponseWriter, r *http.Request) {
	path := r.URL.Path[1:]
	var responseBody []byte

	if r.Method != "GET" {
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		w.WriteHeader(http.StatusMethodNotAllowed)
		w.Header().Set("Allow", "GET")
		responseBody, _ = ioutil.ReadFile("405.min.html")
	} else if len(path) == 0 {
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		responseBody, _ = ioutil.ReadFile("index.html")
	} else if path == `jquery.min.js` {
		w.Header().Set("Content-Type", "application/javascript; charset=utf-8")
		responseBody, _ = ioutil.ReadFile("jquery.min.js")
	} else {
		w.Header().Set("Content-Type", "text/html; charset=utf-8")
		w.WriteHeader(http.StatusNotFound)

		// TODO 以 gzip 壓縮 404 網頁會讓瀏覽器無法正確解析? (Chrome 31.0.1650.63)
		responseBody, _ = ioutil.ReadFile("404.min.html")
		w.Write(responseBody)
		return
	}

	if !strings.Contains(r.Header.Get("Accept-Encoding"), "gzip") {
		w.Write(responseBody)
		return
	}

	w.Header().Set("Content-Encoding", "gzip")
	gz := gzip.NewWriter(w)
	defer gz.Close()

	gz.Write(responseBody)
}

func slideThumbnail(w http.ResponseWriter, r *http.Request) {
	w.Header().Set("Content-Type", "image/png")

	if r.Method != "GET" {
		w.WriteHeader(http.StatusMethodNotAllowed)
		w.Header().Set("Allow", "GET")
		w.Write(_ThumbnailNotFoundImage)
		return
	}

	slideIndex := r.URL.Path[15:]

	if len(slideIndex) < 1 {
		w.WriteHeader(http.StatusBadRequest)
		w.Write(_ThumbnailNotFoundImage)
		return
	}

	filaneme := `presthumb\` + slideIndex + `.png`
	file, err := os.Open(filaneme)
	if err != nil {
		fmt.Printf("file '%s' does not exist\n", filaneme)

		w.WriteHeader(http.StatusNotFound)
		w.Write(_ThumbnailNotFoundImage)
		return
	}

	defer file.Close()

	// make a buffer to keep chunks that are read
	buf := make([]byte, 1024)
	for {
		// read a chunk
		n, err := file.Read(buf)
		if (err != nil && err != io.EOF) || n == 0 {
			break
		}

		// write a chunk
		if _, err := w.Write(buf[:n]); err != nil {
			break
		}
	}
}

func remoteControl(ws *websocket.Conn) {
	pptRequest := &pptctrl.Request{Response: make(chan string)}

	defer ws.Close()
	defer close(pptRequest.Response)

	r := bufio.NewReader(ws)

	for {
		data, err := r.ReadBytes('\n')
		fmt.Printf("%s", data)

		if err != nil {
			fmt.Printf("Error occured: %s\n", err.Error())
			break
		}

		switch data[0] {
		case '!': // PowerPoint control
			pptRequest.Data = data

			pptctrl.SendRequest(pptRequest)

			// block current goroutine
			extraReturnInfo := <-pptRequest.Response

			if len(extraReturnInfo) > 0 {
				sendPowerPointExtraReturnInfo(data, extraReturnInfo, ws)
			}
		}
	}
}

func sendPowerPointExtraReturnInfo(request []byte, extraReturnInfo string, ws *websocket.Conn) {
	switch request[1] {
	case 'p', 'n', 'r', 'l':
		// 傳回目前投影片的備忘稿; 縮圖的網址也直接傳回去，讓瀏覽器判斷
		ws.Write([]byte(extraReturnInfo))
	}

}
