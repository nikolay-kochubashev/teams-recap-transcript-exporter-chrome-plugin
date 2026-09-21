Teams Recap Transcript Exporter v2.0.0

Назначение
----------
Chrome-расширение для экспорта транскрипций Microsoft Teams / SharePoint Recap, в том числе когда Transcript доступен для просмотра, но кнопка Download недоступна.

Режимы
------
1. Текущая встреча
   - Открыть запись встречи / Recap / Transcript.
   - Нажать "Собрать транскрипцию".
   - Экспортировать TXT в Windows Documents.

2. Пакет из календаря
   - Открыть Teams Web -> Calendar.
   - Показать нужный диапазон встреч, например неделю.
   - В Side Panel открыть "Пакет из календаря".
   - Нажать "Считать календарь".
   - Выбрать встречи.
   - Нажать "Собрать выбранные".

Пакетный workflow
-----------------
Для каждой выбранной встречи расширение пытается последовательно:

Calendar -> Meeting details -> Recap/Recording -> Meeting Chat fallback -> SharePoint/Stream Recording -> Transcript.

Каждая транскрипция собирается существующим lazy-loading collector-ом и сохраняется в отдельный TXT.
Ошибка одной встречи не останавливает весь batch.

Результат batch
---------------
Windows Documents\Teams Transcripts\YYYYMMDD-HHmmss\

Внутри создаются:
- <Meeting title> - YYYYMMDD.txt
- <Meeting title> - YYYYMMDD_02.txt, если у встречи несколько записей
- batch-report.txt

Статусы batch-report.txt:
- DONE - транскрипция сохранена
- SKIP - запись/Recap не найдены
- ERROR - встреча найдена, но обработать ее не удалось

Установка расширения
--------------------
1. Открыть chrome://extensions
2. Включить Developer mode.
3. Нажать Load unpacked.
4. Выбрать папку проекта.
5. После обновления расширения обновить вкладки Teams/SharePoint.

Windows helper
--------------
Для пакетного режима Windows helper обязателен.

После перехода с v1.x на v2.0 один раз повторно запустить:

Install-Windows-Integration.cmd

Helper остается on-demand: Chrome запускает его только на время файловой операции.
Он не устанавливается как Windows Service и не висит постоянно в фоне.

Диагностика календаря
---------------------
Если встречи не определяются или автоматический переход в Recap не работает:
1. Открыть Teams Calendar в проблемном состоянии.
2. Нажать "Диагностика календаря".
3. Передать diagnostic TXT для адаптации DOM-эвристик.

Техническая архитектура
-----------------------
- Manifest V3 Chrome Extension
- content.js - существующий SharePoint/Stream Transcript collector
- calendar.js - DOM scanner и UI automation Teams Calendar / Meeting Chat
- batch-background.js - batch coordinator и state machine
- batch-panel.js - UI пакетного режима
- NativeHost.cs - безопасное сохранение в Windows Documents
- chrome.storage.local - состояние пакетного запуска

Ограничение v2.0
----------------
Teams Web не предоставляет стабильный публичный DOM-контракт для Calendar/Meeting Chat. Поэтому навигация сделана через семантические признаки UI и может потребовать адаптации при изменении интерфейса Microsoft. Для этого предусмотрена отдельная диагностика календаря.
