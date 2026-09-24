Teams Recap Transcript Exporter v2.5.3

Назначение
----------
Chrome-расширение для сбора рабочих материалов из Microsoft Teams:
- транскрипции Teams / SharePoint Recap;
- пакетный экспорт транскрипций встреч из Teams Calendar;
- переписки Teams за предыдущую календарную неделю с полным контекстом сообщений участников для подготовки отчета о проделанной работе.

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
   - Указать полное имя пользователя в Teams.
   - Нажать "Собрать переписку".
   - Расширение автоматически выбирает предыдущую календарную неделю: понедельник -> воскресенье.
   - Search используется только для обнаружения релевантных чатов.
   - После этого расширение открывает каждый найденный чат и собирает сообщения всех участников за выбранный период.
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
Режим "Чаты за неделю" использует штатный Teams Search и DOM Teams Web, без Microsoft Graph API.

Discovery выполняется одним точным KQL-запросом:

from:"<полное имя пользователя>"

Далее расширение переключается на Messages, применяет Date filter за предыдущую календарную неделю и обходит все страницы результатов.

Основной маршрут:
Teams Search -> KQL from:"<имя>" -> Messages -> Date предыдущей недели -> все страницы Search Results -> уникальные чаты -> открыть каждый чат -> lazy-loading истории -> все сообщения участников за период -> TXT.

Teams Search используется только для discovery:
- Search result определяет релевантный чат;
- текст Search result не считается полным контекстом;
- в итоговый TXT попадают сообщения, считанные непосредственно из открытого message pane;
- если конкретный чат открыть не удалось, в отчете фиксируется ошибка для этого чата, но Search snippet не подставляется вместо полного контекста.

Основные DOM-маркеры Search:
- ms-searchux-input;
- search-content;
- search-date-filter;
- messages-tab;
- more-Messages;
- search-card;
- message-app-card-header;
- search-pagination-previous-next.

Для чтения открытого чата используются:
- message-pane-list-viewport;
- chat-pane-item;
- chat-pane-message;
- message-author-name;
- time[datetime];
- data-message-content;
- quoted-reply-card.

История в чате виртуализирована. Расширение прокручивает ее вверх до начала периода, затем вниз до конца периода и собирает lazy-loaded сообщения.

Результат содержит:
- период;
- автора, использованного для discovery;
- количество Search hits;
- список релевантных чатов;
- полный диалог за период с сообщениями всех участников;
- дату/время и автора каждого сообщения;
- quoted reply preview, если он есть;
- найденные URL;
- информацию о чатах, которые не удалось открыть.

Результаты
----------
Транскрипции и недельная переписка сохраняются через Windows helper в Documents\Teams Transcripts.

Примеры:
- <Meeting title> - YYYYMMDD.txt
- <Meeting title> - YYYYMMDD - HHMM-HHMM.txt
- Teams chats - YYYYMMDD-YYYYMMDD.txt
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

NativeHost.cs в версии 2.5.3 не менялся относительно предыдущих рабочих версий helper, поэтому повторная установка helper обычно не требуется.

Диагностика
-----------
Для встреч создается batch-operation-log.txt.

Для чатов создается chat-operation-log.txt. В Side Panel также есть кнопка "Скопировать диагностику".

Диагностика чатов содержит:
- текущую версию;
- stage workflow;
- фактический KQL-запрос;
- наличие Search input;
- состояние Date filter;
- количество search-card;
- количество распарсенных сообщений;
- наличие pagination и состояние Next;
- sample результатов;
- состояние открытого message pane.

Техническая архитектура
-----------------------
- Manifest V3 Chrome Extension
- background.js - lifecycle расширения и tab-scoped Side Panel
- content.js - lazy-loading Transcript collector для Teams / SharePoint frames
- calendar.js - DOM adapter и UI automation Teams Calendar / Meeting / Recap
- batch-background.js - coordinator пакетного экспорта встреч
- batch-panel.js - UI пакетного режима встреч
- chat-search.js - Teams Search adapter, KQL discovery и сбор полного контекста из message pane
- chat-background.js - coordinator discovery чатов и недельного экспорта полного контекста
- chat-panel.js - UI режима "Чаты за неделю"
- NativeHost.cs - сохранение файлов в Windows Documents
- chrome.storage.local - состояния batch/chat и сохраненное имя пользователя
- chrome.storage.session - текущая вкладка-владелец Side Panel

Ограничения
-----------
Teams Web не предоставляет стабильный публичный DOM-контракт для Calendar / Search / Meeting / Recap, поэтому DOM adapters могут потребовать обновления при изменениях интерфейса Microsoft.

Batch и chat workflow используют живой DOM вкладки Teams. Chrome может throttling/freeze неактивные вкладки. Полностью resumable orchestration после freeze/discard вкладки пока не реализован.

Teams Search используется только для discovery. Полный текст переписки читается из открытого чата. Если Search result не открывает совместимый message pane, расширение фиксирует ошибку и не заменяет полный контекст Search snippet.
