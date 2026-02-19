# الدليل الشامل لمشروع Mudabbir (AR)

مشروع Mudabbir هو نسخة مُعاد تسميتها ومُخصصة من المشروع الأصلي `mudabbir`.

## 1) العنوان

- اسم المشروع: `Mudabbir`
- هوية الحزمة/التشغيل: `mudabbir` (قد تبقى هذه التسمية داخليًا ضمن مراجع التطوير والمعمارية)
- إصدار بايثون المطلوب: 3.11+

## 2) نظرة عامة على المشروع

Mudabbir هو وكيل ذكاء اصطناعي ذاتي الاستضافة يعمل محليًا، ويمكن التحكم به عبر Telegram وDiscord وSlack وWhatsApp أو عبر لوحة تحكم ويب.

### ما الذي تغيّر

- إعادة العلامة من `mudabbir` إلى `Mudabbir`.
- إضافة طبقة توافق وقت التشغيل بحيث تعمل استيرادات `Mudabbir` بشكل صحيح على بيئة Windows الحالية.
- تحسين موثوقية الإقلاع:
  - أمان أعلى في التعامل مع مخزن الأسرار وسلوك الترحيل،
  - افتراضات أفضل عند بدء WhatsApp في الوضع الشخصي،
  - تحسينات ثبات في المشغل.
- تحسينات في سلوك الأوامر المباشرة المكتبية ضمن تدفق Open Interpreter.

## 3) البدء السريع

### المسار الرئيسي للمشروع

- `C:\Users\Admin\Documents\Mudabbir`

### أوامر التشغيل

- `Mudabbir --port 8888`
- `python -m Mudabbir --port 8888`

### لوحة الويب

- افتح: `http://127.0.0.1:8888`

### ملفات محلية فقط (غير مخصصة للرفع)

- `C:\Users\Admin\Documents\for Mudabbir\local-only`
- يحتوي هذا المسار على المشغل المحلي، وسكربتات التشغيل المساعدة، والاختبارات، وأدوات النسخ الاحتياطي/الإصدارات.

## 4) حالة المشروع

### المستودع الأساسي

- `C:\Users\Admin\Documents\Mudabbir`

### النسخة المرآتية العاملة

- `C:\Python314\Lib\site-packages\Mudabbir\Mudabbir-main`

### أوامر التشغيل

- `Mudabbir`
- `Mudabbir --port 8888`
- `python -m Mudabbir --port 8888`

### حالة العلامة

- تم حذف نصوص العلامة القديمة Pocket/Paw من الملفات النصية.
- تم إعادة تسمية أصول `paw/pocket` إلى أسماء Mudabbir.

### حالة Git

- الفرع: `main`
- أحدث التعديلات تتضمن إعادة التسمية والتنظيف الكامل.
- توفر ملفات bundle يعتمد على أدوات التشغيل المحلية خارج الريبو.

## 5) التطوير والأوامر

### الأوامر الأساسية

```bash
# تثبيت اعتماديات التطوير
uv sync --dev

# تشغيل التطبيق (لوحة الويب هي الوضع الافتراضي — يبدأ كل الـ adapters المهيأة تلقائيًا)
uv run mudabbir

# تشغيل وضع Telegram فقط (تدفق الربط القديم)
uv run mudabbir --telegram

# تشغيل بوت Discord بدون واجهة
uv run mudabbir --discord

# تشغيل بوت Slack بدون واجهة (Socket Mode بدون رابط عام)
uv run mudabbir --slack

# تشغيل خادم WhatsApp webhook بدون واجهة
uv run mudabbir --whatsapp

# تشغيل عدة قنوات بدون واجهة معًا
uv run mudabbir --discord --slack

# تشغيل في وضع التطوير (إعادة تحميل تلقائي عند تغيّر الملفات)
uv run mudabbir --dev

# تشغيل كل الاختبارات
uv run pytest

# تشغيل ملف اختبار واحد
uv run pytest tests/test_bus.py

# تشغيل اختبار محدد
uv run pytest tests/test_bus.py::test_publish_subscribe -v

# تشغيل الاختبارات بدون e2e (لأنها تحتاج Playwright browsers)
uv run pytest --ignore=tests/e2e

# تشغيل ملف اختبار واحد (بديل)
uv run pytest tests/test_bus.py -v

# فحص lint
uv run ruff check .

# تنسيق
uv run ruff format .

# فحص الأنواع
uv run mypy .

# بناء الحزمة
python -m build
```

## 6) المعمارية

### نمط Message Bus

قلب المعمارية يعتمد على ناقل أحداث event-driven موجود في (`src/mudabbir/bus/`). جميع الاتصالات تمر عبر ثلاثة أنواع أحداث معرفة في `bus/events.py`:

- **InboundMessage** — مدخلات المستخدم من أي قناة (Telegram، WebSocket، CLI)
- **OutboundMessage** — ردود الوكيل إلى القنوات (يدعم البث عبر `is_stream_chunk` و`is_stream_end`)
- **SystemEvent** — أحداث داخلية (tool_start, tool_result, thinking, error) تستهلكها لوحة Activity في الويب

### خط التدفق AgentLoop -> AgentRouter -> Backend

خط المعالجة موجود في `agents/loop.py` و`agents/router.py`:

1. **AgentLoop** يستهلك الرسائل من ناقل الأحداث، يدير سياق الذاكرة، ويرسل الردود بشكل متدفق.
2. **AgentRouter** يختار ويفوض لأحد ثلاثة backends بحسب `settings.agent_backend`:
   - `claude_agent_sdk` (الافتراضي/الموصى به) — Claude Agent SDK الرسمي مع أدوات مدمجة (Bash, Read, Write, ...). يستخدم `PreToolUse` لحجب الأوامر الخطرة. في `agents/claude_sdk.py`. إعدادات خاصة: `claude_sdk_model` (فارغ = Claude Code يختار تلقائيًا)، `claude_sdk_max_turns` (الافتراضي 25). توجيه النماذج الذكي معطّل افتراضيًا لتجنب التعارض مع توجيه Claude Code.
   - `mudabbir_native` — orchestrator مخصص: Anthropic SDK للاستدلال + Open Interpreter للتنفيذ. في `agents/mudabbir_native.py`.
   - `open_interpreter` — Open Interpreter مستقل يدعم Ollama/OpenAI/Anthropic. في `agents/open_interpreter.py`.
3. كل backends تُرجع مخرجات موحدة بصيغة dict تحتوي `type` (`message/tool_use/tool_result/error/done`) و`content` و`metadata`.

### Channel Adapters

المجلد `bus/adapters/` يحتوي مترجمات البروتوكولات التي تربط القنوات الخارجية بناقل الرسائل:

- `TelegramAdapter` — python-telegram-bot
- `WebSocketAdapter` — FastAPI WebSockets
- `DiscordAdapter` — discord.py (اعتمادية اختيارية `mudabbir[discord]`). يدعم أمر `/mudabbir` + الرسائل الخاصة/المنشن. مع buffering وتعديل في نفس الرسالة (rate limit 1.5s).
- `SlackAdapter` — slack-bolt Socket Mode (اعتمادية اختيارية `mudabbir[slack]`). يتعامل مع `app_mention` + الرسائل الخاصة. لا يحتاج URL عام. يدعم threads عبر `thread_ts`.
- `WhatsAppAdapter` — WhatsApp Business Cloud API عبر `httpx` (اعتمادية أساسية). لا يدعم streaming مباشر؛ يجمع chunks ويرسل عند `stream_end`. لوحة الويب توفر مسارات `/webhook/whatsapp`؛ والوضع المنفصل يشغل FastAPI خاصًا به.

إدارة القنوات من لوحة التحكم:

- وضع لوحة الويب (الافتراضي) يبدأ كل adapters المهيأة تلقائيًا.
- يمكن تهيئة/بدء/إيقاف القنوات من نافذة Channels في الشريط الجانبي.
- REST API:
  - `GET /api/channels/status`
  - `POST /api/channels/save`
  - `POST /api/channels/toggle`

### الأنظمة الفرعية الأساسية

- **Memory** (`memory/`) — تاريخ الجلسة + حقائق طويلة الأمد، تخزين ملفي في `~/.mudabbir/memory/`. قائم على `MemoryStoreProtocol` لتبديل backend مستقبلًا.
- **Browser** (`browser/`) — أتمتة Playwright عبر لقطات شجرة الوصول accessibility tree (وليس screenshots). `BrowserDriver` يرجع `NavigationResult` مع `refmap` يربط أرقام refs بـ CSS selectors.
- **Security** (`security/`) — Guardian AI (فحص أمان عبر LLM ثانوي) + سجل تدقيق append-only في `~/.mudabbir/audit.jsonl`.
- **Tools** (`tools/`) — `ToolProtocol` مع `ToolDefinition` يدعم تصدير schema لكل من Anthropic وOpenAI. الأدوات المدمجة ضمن `tools/builtin/`.
- **Bootstrap** (`bootstrap/`) — `AgentContextBuilder` يبني system prompt من الهوية والذاكرة والحالة الحالية.
- **Config** (`config.py`) — إعدادات Pydantic مع بادئة `MUDABBIR_`، وملف JSON في `~/.mudabbir/config.json`. إعدادات قنوات تشمل:
  - `discord_bot_token`
  - `discord_allowed_guild_ids`
  - `discord_allowed_user_ids`
  - `slack_bot_token`
  - `slack_app_token`
  - `slack_allowed_channel_ids`
  - `whatsapp_access_token`
  - `whatsapp_phone_number_id`
  - `whatsapp_verify_token`
  - `whatsapp_allowed_phone_numbers`

### الواجهة الأمامية

لوحة الويب (`frontend/`) مبنية بـ JS/CSS/HTML بدون build step، وتُخدم عبر FastAPI + Jinja2، وتتواصل مع الباكند عبر WebSocket للبث اللحظي.

### بنية المشروع

```text
src/mudabbir/
  agents/            # Backends للوكلاء (Claude SDK, Native, Open Interpreter) + router
  bus/               # Message bus + أنواع الأحداث
    adapters/        # Adapters القنوات (Telegram, Discord, Slack, WhatsApp, ...)
  tools/
    builtin/         # أكثر من 60 أداة مدمجة
    protocol.py      # واجهة ToolProtocol (طبّقها عند إضافة أدوات)
    registry.py      # سجل الأدوات المركزي مع policy filtering
    policy.py        # التحكم بصلاحيات الأدوات
  memory/            # مخازن الذاكرة (file-based, mem0)
  security/          # Guardian AI، فحص الحقن، سجل التدقيق
  mcp/               # إعداد وإدارة MCP server
  deep_work/         # تفكيك وتنفيذ المهام متعددة الخطوات
  mission_control/   # تنسيق متعدد الوكلاء
  daemon/            # مهام خلفية ومحفزات وسلوك استباقي
  config.py          # إعدادات Pydantic مع بادئة MUDABBIR_
  credentials.py     # مخزن بيانات اعتماد مشفر (Fernet)
  dashboard.py       # خادم FastAPI + WebSocket + REST APIs
  scheduler.py       # تذكيرات ومهام دورية عبر APScheduler
frontend/            # لوحة JS/CSS/HTML بدون build step
tests/               # مجموعة pytest (130+ اختبار)
```

### قواعد أساسية

- **Async في كل شيء**: واجهات agent/bus/memory/tools كلها async. الاختبارات تستخدم `pytest-asyncio` مع `asyncio_mode = "auto"`.
- **Protocol-oriented**: الواجهات الأساسية (`AgentProtocol`, `ToolProtocol`, `MemoryStoreProtocol`, `BaseChannelAdapter`) مبنية بـ Python `Protocol` لتسهيل تبديل التنفيذ.
- **متغيرات البيئة**: كل الإعدادات تبدأ بـ `MUDABBIR_` (مثل `MUDABBIR_ANTHROPIC_API_KEY`).
- **Ruff config**: حد السطر 100، Python 3.11، قواعد E/F/I/UP.
- **Entry point**: `mudabbir.__main__:main`.
- **Lazy imports**: يتم استيراد backends داخل `AgentRouter._initialize_agent()` لتجنب تحميل الاعتماديات غير المستخدمة.

## 7) المساهمة

Mudabbir مشروع مفتوح المصدر. نرحب بكل أنواع المساهمات: إصلاحات أخطاء، أدوات جديدة، adapters للقنوات، توثيق، واختبارات.

### استراتيجية الفروع

> **كل Pull Requests يجب أن تستهدف فرع `dev`.**
>
> أي PR على `main` سيتم إغلاقه. فرع `main` يُحدّث فقط عبر merge من `dev` عند الجاهزية للإصدار.

### قبل البدء

- ابحث في القضايا الحالية: <https://github.com/Ahmed5754/Mudabbir/issues>
- راجع PRs المفتوحة: <https://github.com/Ahmed5754/Mudabbir/pulls>
- إذا كانت القضية موجودة، علّق أنك ستعمل عليها.
- إذا لا توجد قضية، افتح واحدة أولًا وناقش النهج.
- نقطة بداية جيدة: <https://github.com/Ahmed5754/Mudabbir/labels/good%20first%20issue>

### إعداد البيئة

1. اعمل Fork للمستودع ثم Clone لنسختك.
2. أنشئ فرع ميزة من `dev`:
   ```bash
   git checkout dev
   git pull origin dev
   git checkout -b feat/your-feature
   ```
3. ثبّت الاعتماديات:
   ```bash
   uv sync --dev
   ```
4. تحقّق من التشغيل:
   ```bash
   uv run mudabbir
   ```
   يجب أن تفتح لوحة الويب على `http://localhost:8888`.

### كتابة الكود

#### القواعد

- Async في كل شيء.
- تصميم protocol-oriented.
- Ruff: line-length 100، Python 3.11، قواعد E/F/I/UP.
- استخدم lazy imports للاعتماديات الاختيارية/الثقيلة.

#### إضافة أداة جديدة

1. أنشئ ملفًا في `src/mudabbir/tools/builtin/`.
2. ورّث من `BaseTool` في `tools/protocol.py`.
3. نفّذ `name`, `description`, `parameters` (JSON Schema), و`execute(**params) -> str`.
4. أضف الصنف إلى lazy imports في `tools/builtin/__init__.py`.
5. سجّل الأداة ضمن policy group المناسبة في `tools/policy.py`.
6. أضف اختبارات.

#### إضافة Channel Adapter جديد

1. أنشئ ملفًا في `src/mudabbir/bus/adapters/`.
2. وسّع `BaseChannelAdapter`.
3. نفّذ `_on_start()`, `_on_stop()`, `send(message)`.
4. استخدم `self._publish_inbound()` لدفع الرسائل الواردة.
5. أضف الاعتماديات الاختيارية ضمن extras في `pyproject.toml`.

### اعتبارات الأمان

- لا تسجّل/تكشف أي أسرار.
- أي حقول إعدادات سرية جديدة يجب إضافتها إلى `SECRET_FIELDS` في `credentials.py`.
- الأدوات التي تنفذ shell يجب أن تحترم فحوصات Guardian AI.
- أي endpoint جديد يحتاج auth middleware.
- اختبر أنماط الحقن عند التعامل مع مدخلات المستخدم.

### رسائل الـ Commit

استخدم Conventional Commits:

```text
feat: add Spotify playback tool
fix: handle empty WebSocket message
docs: update channel adapter guide
refactor: simplify model router thresholds
test: add coverage for injection scanner
```

- اجعل عنوان الرسالة أقل من 72 حرفًا.
- أضف body عند الحاجة للشرح.

### قائمة فحص Pull Request

- [ ] الفرع مبني من `dev`
- [ ] الـ PR يستهدف `dev`
- [ ] الاختبارات ناجحة (`uv run pytest --ignore=tests/e2e`)
- [ ] lint ناجح (`uv run ruff check .`)
- [ ] لا توجد أسرار في diff
- [ ] حقول config الجديدة مضافة في `Settings.save()` dict
- [ ] حقول الأسرار الجديدة مضافة في `SECRET_FIELDS`
- [ ] الأدوات الجديدة مسجلة ضمن policy group صحيح
- [ ] الاعتماديات الاختيارية الجديدة مضافة في extras داخل `pyproject.toml`

### المراجعة

- المساهمات تُراجع من الفريق المشرف، وغالبًا خلال عدة أيام.
- الـ PRs الصغيرة والمركزة تُراجع أسرع.
- إذا مر أسبوع بلا رد، اعمل ping في القضية المرتبطة.

### الإبلاغ عن الأخطاء

اذكر:

- ما المتوقع أن يحدث
- ما حدث فعليًا
- خطوات إعادة الإنتاج
- نظام التشغيل + إصدار بايثون + إصدار Mudabbir (`mudabbir --version`)

### الأسئلة

- افتح Discussion: <https://github.com/Ahmed5754/Mudabbir/discussions>
- أو علّق على القضية المناسبة.

## 8) الترخيص (MIT - نسخة عربية)

ترخيص MIT

حقوق النشر (c) 2026 فريق Mudabbir

يُمنح الإذن، مجانًا، لأي شخص يحصل على نسخة من هذا البرنامج والملفات التوثيقية المرتبطة به ("البرنامج") للتعامل مع البرنامج دون قيود، بما في ذلك، على سبيل المثال لا الحصر، حقوق الاستخدام والنسخ والتعديل والدمج والنشر والتوزيع ومنح الترخيص من الباطن و/أو بيع نسخ من البرنامج، وكذلك السماح للأشخاص الذين يُزوَّدون بالبرنامج بممارسة ذلك، وذلك وفقًا للشروط التالية:

يجب تضمين إشعار حقوق النشر أعلاه وإشعار الإذن هذا في جميع النسخ أو الأجزاء الجوهرية من البرنامج.

يتم توفير البرنامج "كما هو"، دون أي ضمان من أي نوع، صريحًا كان أو ضمنيًا، بما في ذلك على سبيل المثال لا الحصر ضمانات القابلية للتسويق، والملاءمة لغرض معين، وعدم الانتهاك. لا يتحمل المؤلفون أو مالكو حقوق النشر، بأي حال من الأحوال، أي مطالبة أو أضرار أو مسؤولية أخرى، سواء في دعوى تعاقدية أو تقصيرية أو غير ذلك، تنشأ عن البرنامج أو تتعلق به أو باستخدامه أو بأي تعاملات أخرى فيه.

