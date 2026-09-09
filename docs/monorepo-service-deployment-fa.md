# دستورالعمل استقرار یک سرویس مونوریپو

این روش برای مونوریپوهای Nx است که خروجی‌ها را به Imageهای immutable تبدیل می‌کند و
Compose مشترک Project/Environment را به‌عنوان مرجع GitOps نگه می‌دارد. پروژهٔ مبدا
به Dockerfile نیاز ندارد و هیچ خروجی Build روی فایل‌سیستم سرور مقصد mount نمی‌شود.

## ۱. آماده‌سازی Nx

هر application قابل استقرار باید target به نام `build` و `outputPath` معتبر داشته
باشد. قرارداد پیش‌فرض این دستورات را از ریشهٔ Repository اجرا می‌کند:

```bash
pnpm install --frozen-lockfile
node tools/scripts/with-env.cjs production pnpm exec nx run-many -t build --projects=<project> --parallel=3
```

قواعد تشخیص:

- Shell با نام `shell` یا `host`، یا tag برابر `deploy:shell` مشخص می‌شود.
- BFF با نام `bff`، یا tag برابر `deploy:bff` مشخص می‌شود؛ فقط یک BFF پشتیبانی می‌شود.
- سایر applicationهای Buildable به‌عنوان ماژول Static شناخته می‌شوند.
- خروجی BFF باید self-contained، با Node قابل اجرا و روی `0.0.0.0:3000` در دسترس
  باشد. فایل شروع پیش‌فرض `main.js` است.
- Shell روی `/`، BFF مونوریپو روی `/bff/` و ماژول Static روی
  `/<nx-project-name>/` ارائه می‌شود؛ `/api/` منحصراً برای Backend اصلی پروژه
  رزرو است و base path و asset URL پروژه‌ها باید با این قرارداد هماهنگ باشد.

افزونه `/.devops/deployments.yml` را فقط در صورت نبودن ایجاد می‌کند. برای نام‌های
غیرپیش‌فرض، entry BFF یا دستورهای اختصاصی install/build همین فایل را ویرایش کنید.

## ۲. Dockerfile و Registry

در Repository پروژه Dockerfile نسازید. فایل‌های عمومی زیر در SharedTemplates هستند:

- `monorepo/docker/static.Dockerfile`: خروجی Shell/Static و کانفیگ Nginx را در
  Image نهایی قرار می‌دهد.
- `monorepo/docker/bff.Dockerfile`: خروجی BFF و entry آن را در Image نهایی قرار
  می‌دهد.
- `monorepo/package-images.sh`: Image قبلی را hydrate، خروجی‌های جدید را overlay،
  و Imageهای جدید را build/push می‌کند.

فرم **Generate MonoRepo** آدرس Registry و Docker Registry Service Connection را
دریافت می‌کند. Build Agent باید Docker داشته باشد؛ Node/pnpm داخل Image داخلی
`registry.buluttakin.com/node:22-bookworm` اجرا می‌شود. Runtimeهای پایه نیز از
Registry انتخاب‌شده خوانده می‌شوند؛ پیش‌فرض‌ها:

- `registry.buluttakin.com/nginx:1.27-alpine`
- `registry.buluttakin.com/node:20-alpine`

Corepack و pnpm داخل همان Container به‌طور پیش‌فرض از Nexus گروهی داخلی زیر
استفاده می‌کنند و مستقیماً به `registry.npmjs.org` متصل نمی‌شوند:

```text
https://registry.buluttakin.com/repository/npm-group
```

پارامتر `npmRegistry` قالب فقط در صورت نیاز به یک npm group امن دیگر override
می‌شود. Store مربوط به pnpm و cache مربوط به Corepack روی مسیر cache عامل
self-hosted mount می‌شوند و نیازی به ساخت یا تغییر `.npmrc` در پروژه نیست.

نام Image خروجی به‌صورت زیر است:

```text
<registry>/<project-key>/<service-key>-<environment>:1.0.<BuildId>
<registry>/<project-key>/<service-key>-bff-<environment>:1.0.<BuildId>
```

## ۳. Compose بدون Host mount

افزونه دو Service مدیریت‌شده را داخل Compose مشترک اضافه می‌کند. نمونهٔ اولیه:

```yaml
services:
  locanit_frontend_demo:
    container_name: locanit_frontend_demo
    image: registry.buluttakin.com/locanit/frontend-demo:${frontend}
    restart: unless-stopped
    expose:
      - "80"
    networks:
      - nginx-network

  locanit_frontend_bff_demo:
    container_name: locanit_frontend_bff_demo
    image: registry.buluttakin.com/locanit/frontend-bff-demo:${frontend_bff}
    profiles: ["mr-frontend-bff"]
    restart: unless-stopped
    expose:
      - "3000"
    networks:
      - nginx-network

networks:
  nginx-network:
    name: nginx-net
    external: true
```

فایل `demo_locanit/.env` کنار Compose در Git نگهداری می‌شود:

```text
frontend:CHANGE_ME
frontend_bff:CHANGE_ME
```

اولین Release مقدارهای `CHANGE_ME` را با tagهای immutable ساخته‌شده جایگزین
می‌کند و از آن پس `image:`های Compose ثابت می‌مانند. اجرای دوبارهٔ Generator:

- Serviceهای نامرتبط و تنظیمات اپراتور را حفظ می‌کند؛
- tag immutable فعال Service مدیریت‌شده را overwrite نمی‌کند؛
- bind mountهای قدیمی `/mnt/graid/projects` و `/var/data/projects`، و همچنین
  `working_dir`/`command` قدیمی BFF را از Serviceهای مدیریت‌شده حذف می‌کند؛
- imageهای bare قدیمی را با پیشوند Registry اصلاح می‌کند.

پیش از هر write، افزونه با `ListDockerNetworks` شبکه‌های واقعی Server انتخابی
Komodo را می‌خواند و فقط نام دقیق `nginx-network` یا `nginx-net` را می‌پذیرد.
کلید منطقی موجود در Compose حفظ می‌شود و فیلد `name` آن را به نام واقعی شبکهٔ
Host متصل می‌کند؛ در نمونهٔ بالا کلید `nginx-network` به شبکهٔ واقعی
`nginx-net` نگاشت شده است. اگر هیچ‌کدام وجود نداشته باشد Generator پیش از
تغییر Git متوقف می‌شود. Nginx بیرونی `/bff/` را بدون rewrite به BFF می‌فرستد،
`/api/` را برای Backend اصلی دست‌نخورده نگه می‌دارد و Location ریشه را در انتها
به Static Runtime می‌فرستد.

در صورت وجود BFF، Release مقدار Profile مدیریت‌شده را از طریق
`COMPOSE_PROFILES` در `.env` فعال می‌کند. Pipeline هر `--profile` قدیمی را از
`extra_args` مربوط به Stack حذف می‌کند، زیرا Komodo آن آرگومان‌ها را بعد از
`docker compose up -d` قرار می‌دهد و Compose آن ترتیب را با خطای
`unknown flag: --profile` رد می‌کند.

## ۴. Build و حفظ خروجی قبلی

Pipeline در شروع Compose فعال را از `<Project>_Docker_DevOps@main` می‌خواند و فقط
Imageای را previous می‌پذیرد که متعلق به Repository مدیریت‌شده باشد و از Registry
pull شود.

- اگر previous معتبر نباشد، همهٔ applicationها Build می‌شوند و baseline ساخته
  می‌شود.
- در اجراهای بعدی Nx فقط applicationهای affected را انتخاب می‌کند؛ تغییر Shell
  همچنان full build است.
- Static Image قبلی extract می‌شود؛ خروجی‌های موفق روی آن overlay می‌شوند و
  خروجی‌های unaffected یا ناموفق قبلی باقی می‌مانند.
- اگر BFF موفق و affected باشد Image جدید می‌گیرد؛ در غیر این صورت tag BFF قبلی
  حفظ می‌شود.
- شکست Shell همیشه Build را متوقف می‌کند. شکست ماژول معمولی بدون نسخهٔ قبلی نیز
  خطاست. در حضور نسخهٔ قبلی، Build با `SucceededWithIssues` ادامه می‌یابد.

Artifact به نام `mr-drop` شامل manifest schema 2، inventory، خروجی‌های موفق و
ارجاع دقیق Imageهای جدید/قبلی است. Stack در مرحلهٔ Build deploy نمی‌شود؛ فقط Repo،
Stack و Profile اختیاری BFF reconcile می‌شوند.

## ۵. Release و Rollback

Release به Node، Docker، دسترسی filesystem سرور یا Komodo Terminal نیاز ندارد:

1. `manifest.json` را از Artifact می‌خواند.
2. Repository داکر را روی `main` clone می‌کند.
3. `compose.yml` را روی Repository ثابت و متغیر نگه می‌دارد؛ برای نمونه
   `image: .../front-monorepo-demo:${front_monorepo}`. سپس tag دقیق Static و در
   صورت وجود BFF را در فایل `.env` همان دایرکتوری با کلیدهای
   `front_monorepo` و `front_monorepo_bff` تغییر می‌دهد و commit/push می‌کند.
   اولین Release جدید، Compose قدیمی دارای tag هاردکد را همراه با `.env` در یک
   commit مهاجرت می‌دهد؛ Releaseهای بعدی فقط `.env` را تغییر می‌دهند.
4. `DeployStack` را اجرا و `GetUpdate` را تا `Complete/success=true` poll می‌کند.
5. اگر deploy شکست بخورد، نسخه دقیق قبلی `compose.yml` و `.env` را با commit
   rollback برمی‌گرداند و Stack را دوباره deploy می‌کند؛ خود Release همچنان
   قرمز می‌شود تا خطای واقعی پنهان نشود.

اگر Komodo هنگام `CreateStack`، `UpdateStack` یا `DeployStack` پاسخ موقت
`Stack busy` بدهد، عملیات حداکثر در پنج تلاش و با فاصلهٔ پنج ثانیه تکرار
می‌شود. خطاهای دسترسی، اعتبارسنجی و سایر خطاهای دائمی بدون retry متوقف می‌شوند.
این مقادیر با `KOMODO_STACK_BUSY_MAX_ATTEMPTS` و
`KOMODO_STACK_BUSY_RETRY_SECONDS` قابل تنظیم‌اند.

بنابراین مسیرهای `/mnt/graid/projects` و `/var/data/projects` در اجرای مونوریپو
نقشی ندارند. `projects_root` مرکزی می‌تواند برای workflowهای قدیمی باقی بماند،
اما Compose و Release جدید از آن استفاده نمی‌کنند.

## ۶. دسترسی‌های لازم

- Docker Registry Service Connection: مجوز login/pull/push برای Build Pipeline.
- `AZP_TOKEN`: خواندن Compose در Build و read/write Repository داکر در Release.
- `KOMODO_API_KEY` و `KOMODO_API_SECRET`: خواندن/ایجاد/به‌روزرسانی Repo و Stack و
  اجرای `DeployStack`/`GetUpdate`.
- Terminal permission روی Komodo لازم نیست.

## ۷. ترتیب تست اولیه

1. افزونهٔ جدید را نصب و صفحهٔ Azure DevOps را hard refresh کنید.
2. **Generate MonoRepo** را روی Branch موردنظر و Environment آزمایشی اجرا کنید.
3. لینک‌های contract، Pipeline، Compose و Nginx را بررسی کنید؛ Compose نباید
   `volumes:`, `working_dir:` یا مسیر host داشته باشد.
4. Pipeline را دستی اجرا کنید. اجرای اول باید full build باشد و Imageهای
   `1.0.<BuildId>` را push کند.
5. Artifact manifest را بررسی کنید: `deploymentMode` باید `immutable-images` و
   `staticImage` باید tag همان Build باشد.
6. Release همان Build را اجرا کنید؛ سپس commit Compose، نتیجهٔ DeployStack و
   container imageهای فعال را بررسی کنید.
7. یک ماژول Static را تغییر دهید و Pipeline را دوباره اجرا کنید؛ لاگ باید previous
   image را پیدا کند و فقط affectedها را Build کند.
8. برای تست rollback، فقط در Environment آزمایشی یک deploy نامعتبر کنترل‌شده
   انجام دهید و تأیید کنید tagهای قبلی به Compose برگشته‌اند.

## ۸. فایل‌های مرکزی

فایل‌های لازم در `ShonizCollection/SharedTemplates/SharedTemplates:/monorepo`:

- `pipeline.yml`
- `mr-build.cjs`
- `package-images.sh`
- `docker/static.Dockerfile`
- `docker/bff.Dockerfile`
- `nginx/default.conf`

Pipeline کوچک پروژه فقط Repositoryهای `SharedTemplatesRepo` و `sourceRepo` و
پارامترهای محیط/Registry را تعریف می‌کند؛ Dockerfile پروژه‌ای لازم نیست.
