Teams Recap Transcript Exporter v2.1.0

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
   - Открыть Side Panel расширения на этой вкладке.
   - Нажать "Считать календарь".
   - Выбрать встречи.
   - Нажать "Собрать выбранные".

Side Panel
----------
Начиная с v2.1.0 Side Panel привязан к конкретной вкладке Chrome.

Как это работает:
- нажать иконку расширения на нужной вкладке Teams;
- Side Panel откроется только для этой вкладки;
- при переключении на другую вкладку панель скрывается;
- при возврате на рабочую вкладку Teams панель снова доступна;
- если открыть расширение на другой вкладке того же окна, рабочая вкладка Side Panel переносится туда.

Это устраняет ситуацию, когда панель расширения постоянно отображается на всех вкладках одного окна Chrome.

Пакетный workflow
-----------------
Расширение считывает встречи из видимого диапазона Teams Calendar и для каждой выбранной встречи связывает Calendar occurrence с соответствующим Recap/Transcript.

Основной маршрут:
Calendar -> конкретная occurrence -> Meeting / Recap -> Transcript -> lazy-loading collector -> TXT.

Для повторяющихся встреч учитывается конкретная дата и временное окно встречи из Calendar.

Если внутри одного календарного слота было несколько фактических сессий, например встречу несколько раз запускали и останавливали, экспортируются все Recap/Transcript, пересекающиеся с этим временным окном.

Если внутри одной фактической сессии несколько Transcript, экспортируются все найденные Transcript.

Ошибка одной встречи не останавливает весь batch.

Результат batch
---------------
Windows Documents\Teams Transcripts\YYYYMMDD-HHmmss\

Примеры файлов:
- <Meeting title> - YYYYMMDD.txt
- <Meeting title> - YYYYMMDD - HHMM-HHMM.txt
- <Meeting title> - YYYYMMDD - HHMM-HHMM_02.txt
- batch-report.txt
- batch-operation-log.txt

Статусы:
- DONE - все найденные транскрипции встречи сохранены
- SKIP - подходящий Recap/Transcript не найден
- ERROR - встреча найдена, но обработать ее не удалось

Установка расширения
--------------------
1. Открыть chrome://extensions
2. Включить Developer mode.
3. Нажать Load unpacked.
4. Выбрать папку проекта.
5. После обновления расширения нажать "Обновить" на chrome://extensions.
6. Обновить вкладки Teams/SharePoint.

После изменения manifest.json вкладки Teams/SharePoint обязательно нужно перезагрузить.

Windows helper
--------------
Для пакетного режима Windows helper обязателен.

После перехода с v1.x на v2.x один раз повторно запустить:

Install-Windows-Integration.cmd

Helper работает on-demand: Chrome запускает его только на время файловой операции.
Он не устанавливается как Windows Service и не висит постоянно в фоне.

Диагностика
-----------
При пакетном запуске автоматически создается batch-operation-log.txt.

Лог содержит:
- распознанную Calendar occurrence;
- дату и временное окно встречи;
- найденные Recap-сессии;
- выбранные Transcript;
- frame, в котором реально загружена транскрипция;
- результаты навигации и fallback;
- ошибки по каждому шагу.

Кнопка "Диагностика календаря" используется, если встречи не определяются или изменилась DOM-разметка Teams.

Техническая архитектура
-----------------------
- Manifest V3 Chrome Extension
- background.js - lifecycle расширения и tab-scoped Side Panel
- content.js - lazy-loading Transcript collector для Teams / SharePoint frames
- calendar.js - DOM adapter и UI automation Teams Calendar / Meeting / Recap
- batch-background.js - batch coordinator и state machine
- batch-panel.js - UI пакетного режима
- NativeHost.cs - сохранение файлов в Windows Documents
- chrome.storage.local - состояние пакетного запуска
- chrome.storage.session - текущая вкладка-владелец Side Panel

Ограничения
-----------
Teams Web не предоставляет стабильный публичный DOM-контракт для Calendar / Meeting / Recap, поэтому адаптер может потребовать обновления при изменениях интерфейса Microsoft.

Текущий batch использует живой DOM вкладки Teams. Chrome может throttling/freeze неактивные вкладки. Переключение на другие вкладки обычно допустимо, но пока не гарантируется полностью resumable-выполнение после freeze/discard вкладки. Архитектура с checkpoint/retry orchestration является отдельным следующим этапом.
