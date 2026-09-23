Teams Recap Transcript Exporter v2.2.0

Назначение
----------
Chrome-расширение для сбора рабочих материалов из Microsoft Teams:
- транскрипции Teams / SharePoint Recap;
- пакетный экспорт транскрипций встреч из Teams Calendar;
- собственные сообщения из Teams Search за предыдущую календарную неделю для подготовки отчета о проделанной работе.

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

3. Чаты за неделю
   - Открыть Teams Web.
   - Открыть Side Panel -> "Чаты за неделю".
   - Указать имя или фамилию, по которой Teams находит текущего пользователя в People. Значение запоминается.
   - Нажать "Собрать переписку".
   - Расширение автоматически выбирает предыдущую календарную неделю: понедельник -> воскресенье.
   - Результат сохраняется в отдельную папку Windows Documents\Teams Transcripts\<timestamp>.

Side Panel
----------
Side Panel привязан к конкретной вкладке Chrome.

Как это работает:
- нажать иконку расширения на нужной вкладке Teams;
- Side Panel откроется только для этой вкладки;
- при переключении на другую вкладку панель скрывается;
- при возврате на рабочую вкладку Teams панель снова доступна;
- если открыть расширение на другой вкладке того же окна, рабочая вкладка Side Panel переносится туда.

Пакетный workflow встреч
------------------------
Расширение считывает встречи из видимого диапазона Teams Calendar и для каждой выбранной встречи связывает Calendar occurrence с соответствующим Recap/Transcript.

Основной маршрут:
Calendar -> конкретная occurrence -> Meeting / Recap -> Transcript -> lazy-loading collector -> TXT.

Для повторяющихся встреч учитывается конкретная дата и временное окно встречи из Calendar.

Если внутри одного календарного слота было несколько фактических сессий, экспортируются все Recap/Transcript, пересекающиеся с этим временным окном.

Если внутри одной фактической сессии несколько Transcript, экспортируются все найденные Transcript.

Ошибка одной встречи не останавливает весь batch.

Workflow чатов
--------------
Режим "Чаты за неделю" использует штатный Teams Search, а не прямой Microsoft Graph API.

Маршрут:
Teams Search -> People / From текущего пользователя -> Date предыдущей недели -> Messages -> все страницы результатов -> TXT.

Используемые стабильные DOM-маркеры Teams:
- AUTOSUGGEST_INPUT;
- AUTOSUGGEST_ACTION_PEOPLECENTRICSEARCH;
- search-people-filter;
- search-date-filter;
- messages-tab;
- more-Messages;
- search-card;
- message-app-card-header;
- search-pagination-previous-next.

Сообщения дедуплицируются по message id, если он доступен, иначе по содержимому карточки.

Результат содержит:
- период;
- автора;
- количество сообщений;
- количество чатов/каналов;
- список чатов/каналов;
- для каждого результата время, чат/канал, текст сообщения и найденные URL.

Важно: Teams Search иногда возвращает длинное сообщение как сокращенный preview с "...". Версия 2.2.0 сохраняет именно текст, который отдает Search result card, без открытия каждого сообщения в исходном чате.

Результаты
----------
Транскрипции и недельная переписка сохраняются через Windows helper в Documents\Teams Transcripts.

Примеры:
- <Meeting title> - YYYYMMDD.txt
- <Meeting title> - YYYYMMDD - HHMM-HHMM.txt
- Teams messages - YYYYMMDD-YYYYMMDD.txt
- batch-report.txt
- batch-operation-log.txt
- chat-operation-log.txt

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
Для пакетного режима встреч и режима чатов Windows helper обязателен.

Helper работает on-demand: Chrome запускает его только на время файловой операции.
Он не устанавливается как Windows Service и не висит постоянно в фоне.

Для v2.2.0 NativeHost.cs не менялся, поэтому повторная установка helper после v2.1.x не требуется.

Диагностика
-----------
Для встреч создается batch-operation-log.txt.

Для чатов создается chat-operation-log.txt. В Side Panel также есть кнопка "Скопировать диагностику", которая собирает состояние Teams Search DOM:
- найден ли Search input;
- From filter;
- Date filter;
- количество search-card;
- распарсенные сообщения;
- наличие pagination;
- состояние Next;
- sample результатов.

Техническая архитектура
-----------------------
- Manifest V3 Chrome Extension
- background.js - lifecycle расширения и tab-scoped Side Panel
- content.js - lazy-loading Transcript collector для Teams / SharePoint frames
- calendar.js - DOM adapter и UI automation Teams Calendar / Meeting / Recap
- batch-background.js - coordinator пакетного экспорта встреч
- batch-panel.js - UI пакетного режима встреч
- chat-search.js - DOM adapter Teams Search
- chat-background.js - coordinator недельного экспорта сообщений
- chat-panel.js - UI режима "Чаты за неделю"
- NativeHost.cs - сохранение файлов в Windows Documents
- chrome.storage.local - состояния batch/chat и сохраненное имя пользователя
- chrome.storage.session - текущая вкладка-владелец Side Panel

Ограничения
-----------
Teams Web не предоставляет стабильный публичный DOM-контракт для Calendar / Search / Meeting / Recap, поэтому DOM adapters могут потребовать обновления при изменениях интерфейса Microsoft.

Текущие batch workflow используют живой DOM вкладки Teams. Chrome может throttling/freeze неактивные вкладки. Полностью resumable orchestration после freeze/discard вкладки пока не реализован.

Teams Search может обрезать длинные сообщения в карточке результата. Для подготовки отчета обычно достаточно поискового preview; получение полного текста каждого длинного сообщения потребует отдельного перехода к исходному сообщению.
