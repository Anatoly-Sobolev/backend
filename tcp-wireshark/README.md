# TCP и Wireshark

Практическая работа по анализу TCP-соединения при выполнении HTTPS-запроса к `example.com`.

Студент: Соболев Анатолий  
Группа: ПИН252т

## Состав работы

- `Отчёт_TCP_Wireshark_Соболев_ПИН252т.docx` — итоговый отчёт Word;
- `sources/capture-example-com.pcapng` — исходный дамп сетевого трафика;
- `screenshots/01_tcp_handshake.png` — трёхэтапное установление соединения;
- `screenshots/02_tcp_close.png` — четырёхэтапное завершение соединения;
- `screenshots/03_chrome_waterfall.png` — этапы запроса в Chrome DevTools.

## Команда запроса

```powershell
curl.exe --interface 192.168.0.10 --http1.1 -H "Connection: close" -o NUL https://example.com
```

## Основные результаты

- 3-way handshake: кадры 1, 2, 3;
- 4-way close: кадры 13, 17, 19, 20;
- HTTP-запрос внутри TLS передан одним TCP-сегментом, кадр 11;
- общее время запроса в Chrome DevTools: 164,18 мс.
