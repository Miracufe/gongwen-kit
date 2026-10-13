banner.png 的源。重建：

    chrome --headless --no-proxy-server --hide-scrollbars --force-device-scale-factor=2 \
      --window-size=1280,400 --screenshot=../banner.png "file://$(pwd)/banner.html"

samples.png：notice、qingshi、xinhan、jiyao 四份样张经 Word 导出 PDF 后的首页（90 dpi）拼图。
出图前设 `GONGWEN_FONT_TITLE="Songti SC Black"`（本机没装方正小标宋时），否则标题回退成宋体。
