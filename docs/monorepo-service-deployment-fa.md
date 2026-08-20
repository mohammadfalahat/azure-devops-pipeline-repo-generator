# دستورالعمل استقرار یک سرویس مونوریپو

این روش برای مونوریپوهای Nx است که می‌خواهند فقط ماژول‌های تغییرکرده را Build و
Deploy کنند، بدون آنکه برای هر ماژول Dockerfile جدا یا Image جدید ساخته شود.

## ۱. آماده‌سازی مونوریپو

هر پروژهٔ قابل استقرار باید در Nx یک target به نام `build` و یک `outputPath`
معتبر داشته باشد. حداقل دستورات پروژه باید از ریشهٔ Repository اجرا شوند:

```bash
pnpm install --frozen-lockfile
node tools/scripts/with-env.cjs production pnpm exec nx run-many -t build --projects=<project> --parallel=3
```

قواعد تشخیص ماژول‌ها:

- Shell با نام `shell` یا `host`، یا tag برابر `deploy:shell` مشخص می‌شود.
- BFF با نام `bff`، یا tag برابر `deploy:bff` مشخص می‌شود؛ فقط یک BFF پشتیبانی می‌شود.
- سایر پروژه‌های Buildable به‌عنوان ماژول Static شناخته می‌شوند.
- BFF باید روی `0.0.0.0:3000` گوش کند و خروجی آن مستقیماً با Node قابل اجرا باشد؛
  پیش‌فرض فایل شروع `main.js` است.
- Shell از `/`، BFF از `/api/` و هر ماژول Static از
  `/<nx-project-name>/` ارائه می‌شود. تنظیمات base path و asset URL فرانت باید با
  این مسیرها سازگار باشد.

افزونه فایل `/.devops/deployments.yml` را خودکار ایجاد می‌کند. در حالت معمول
نیازی به تغییر آن نیست. فقط برای نام متفاوت Shell/BFF، فایل شروع BFF یا دستور
Build اختصاصی آن را ویرایش کنید.

## ۲. Dockerfile

در این روش پروژهٔ مونوریپو به Dockerfile نیاز ندارد. Pipeline کد را روی Build
Agent کامپایل می‌کند و Release فقط خروجی Build را داخل Runtimeهای ثابت زیر قرار
می‌دهد:

- `nginx:1.27-alpine` برای Shell و ماژول‌های Static؛
- `node:20-alpine` برای BFF.

بنابراین تغییر کد یا اضافه‌شدن ماژول باعث `docker build` نمی‌شود. اگر استفاده
از Image عمومی مجاز نیست، تیم زیرساخت باید یک‌بار معادل همین دو Runtime را در
Registry داخلی بسازد و فقط مقدار `image` در Compose را تغییر دهد. Dockerfile
چنین Runtimeهایی متعلق به Repository زیرساخت است، نه تک‌تک ماژول‌های مونوریپو.

## ۳. فایل Compose

افزونه Compose مستقلی برای مونوریپو نمی‌سازد؛ سرویس مونوریپو را داخل همان
`compose.yml` مشترک Project و Environment قرار می‌دهد. نمونهٔ بخش افزوده‌شده
برای Repository سرویس `frontend` در `Locanit / dev` به این صورت است:

```yaml
services:
  locanit_frontend_dev:
    container_name: locanit_frontend_dev
    image: nginx:1.27-alpine
    restart: unless-stopped
    volumes:
      - /mnt/graid/projects/Locanit_Docker_DevOps/dev_locanit/monorepo/frontend:/srv/monorepo:ro
      - /mnt/graid/projects/Locanit_Docker_DevOps/dev_locanit/monorepo/frontend/runtime/nginx/default.conf:/etc/nginx/conf.d/default.conf:ro
    expose:
      - "80"
    networks:
      - nginx-network

  locanit_frontend_bff_dev:
    container_name: locanit_frontend_bff_dev
    image: node:20-alpine
    profiles: ["mr-frontend-bff"]
    restart: unless-stopped
    working_dir: /srv/monorepo/current/modules/${MR_LOCANIT_FRONTEND_DEV_BFF_PROJECT:-bff}
    command: ["/bin/sh", "-ec", "exec node \"${MR_LOCANIT_FRONTEND_DEV_BFF_ENTRY:-main.js}\""]
    volumes:
      - /mnt/graid/projects/Locanit_Docker_DevOps/dev_locanit/monorepo/frontend:/srv/monorepo:ro
    expose:
      - "3000"
    networks:
      - nginx-network

networks:
  nginx-network:
    external: true
```

از دید استقرار، Repository مونوریپو یک سرویس منطقی داخل همان Stack است. ورودی
اول Runtime اصلی آن است و فقط اگر Nx یک BFF پیدا کند، ورودی دوم به‌عنوان
companion همان سرویس با Profile اختصاصی فعال می‌شود؛ هیچ Stack مستقلی ساخته
نمی‌شود.

نکات Compose:

- این بخش در همان
  `Locanit_Docker_DevOps:/dev_locanit/compose.yml@main` کنار سایر سرویس‌ها قرار
  می‌گیرد و همان فایل مرجع GitOps کل Stack است؛ Compose یا Stack جداگانه‌ای با
  پیشوند `MR` ساخته نمی‌شود.
- اگر `compose.yml` از قبل وجود داشته باشد، افزونه فقط Service entryهای غایب را
  اضافه می‌کند و سرویس‌های موجود یا ویرایش‌های اپراتور را تغییر نمی‌دهد.
- نام Containerها باید با upstreamهای Nginx بیرونی یکسان بماند.
- مسیر Host باید به ریشهٔ مدیریت‌شدهٔ همان Project و Environment اشاره کند؛
  Release پوشه‌های نسخه‌ای و symlink به نام `current` را در همین مسیر می‌سازد.
- Volume باید فقط‌خواندنی باشد و کل ریشهٔ Monorepo را mount کند، نه یک Build
  مشخص را؛ با تغییر `current` نسخهٔ فعال بدون تعویض Image عوض می‌شود.
- کانفیگ مرکزی Nginx از Artifact در
  `<MR_DEPLOY_ROOT>/runtime/nginx/default.conf` قرار می‌گیرد و به‌صورت
  فقط‌خواندنی روی `/etc/nginx/conf.d/default.conf` mount می‌شود.
- `nginx-network` باید از قبل به‌صورت external وجود داشته باشد. به‌جای `ports`
  از `expose` استفاده می‌شود چون Nginx بیرونی روی همین Network متصل است.
- Profile مربوط به BFF فقط وقتی فعال می‌شود که خروجی BFF در Release موجود باشد.

## ۴. ایجاد و اجرای استقرار

1. از منوی Branch گزینهٔ **Generate MonoRepo** را اجرا و Environment و Komodo
   Server را انتخاب کنید.
2. افزونه Repositoryهای Azure/Docker/Nginx DevOps، فایل قرارداد، Pipeline و
   Release با پیشوند `MR` را ایجاد یا همگام می‌کند، اما Runtime مونوریپو را در
   Compose و Stack معمولی همان Project/Environment merge می‌کند.
3. لینک‌های پایان فرم را باز کنید و به‌ترتیب `deployments.yml`، `compose.yml` و
   کانفیگ Nginx را بررسی کنید.
4. Pipeline را بار اول دستی اجرا کنید. Pipeline پروژه‌های Buildable و affected
   را از Nx می‌گیرد، هر ماژول را جداگانه Build می‌کند و همان Komodo Repo/Stack
   معمولی متصل به `compose.yml` مشترک ADO را به‌صورت partial ایجاد یا
   به‌روزرسانی می‌کند؛ تنظیمات سایر سرویس‌ها حفظ و Stack در Build اجرا نمی‌شود.
5. Pipeline یک Artifact به نام `mr-drop` شامل inventory کامل، خروجی ماژول‌های
   موفق و `runtime/nginx/default.conf` مرکزی می‌سازد. سپس Release مربوط به همان
   Build را اجرا کنید.
6. Release از طریق Terminal API نسخهٔ `1.19.5` کومودو Artifact را روی سرور
   آماده می‌کند، نسخهٔ جدید را کنار نسخهٔ قبلی می‌سازد و symlink `current` را
   اتمیک جابه‌جا می‌کند. سپس `DeployStack` را برای Stack متصل به Git اجرا و
   نتیجهٔ async را با `GetUpdate` تا `Complete/success=true` پیگیری می‌کند و پس
   از `nginx -t` کانفیگ Container استاتیک را reload می‌کند. در شکست، symlink و
   کانفیگ قبلی برگردانده و Stack دوباره روی وضعیت قبلی Deploy می‌شود.

در تغییرات بعدی فقط پروژه‌های affected ساخته می‌شوند. شکست یک ماژول معمولی
نسخهٔ قبلی همان ماژول را حفظ می‌کند و مانع استقرار ماژول‌های موفق نمی‌شود؛ شکست
Shell کل استقرار را متوقف می‌کند. ماژول جدید خودکار کشف می‌شود و ماژول حذف‌شده
یا تغییرنام‌یافته برای بررسی دستی نگه داشته می‌شود و خودکار پاک نمی‌شود.

## ۵. بررسی نهایی

- همهٔ پروژه‌های قابل استقرار `build` و `outputPath` صحیح دارند.
- Shell و BFF با نام قراردادی یا tag مناسب مشخص شده‌اند.
- خروجی BFF مستقل، قابل اجرای مستقیم با Node و در دسترس روی پورت 3000 است.
- base path ماژول‌های فرانت با نام Nx آنها هماهنگ است.
- نام Container، مسیر Host، Network و Imageهای Compose بررسی شده‌اند.
- Komodo Repo و Stack معمولی Project/Environment به Repository داکر ADO روی
  `main` و مسیر دقیق `compose.yml` مشترک متصل هستند؛ Stack جداگانهٔ MR وجود ندارد.
- کانفیگ Nginx بیرونی `/api/` را بدون rewrite به BFF و `/` را در آخر به Runtime
  Static می‌فرستد.

## ۶. قالب مرکزی SharedTemplates

برای استفادهٔ همهٔ Pipelineهای مونوریپو از یک پیاده‌سازی، این سه فایل را در
`ShonizCollection/SharedTemplates/SharedTemplates:/monorepo` قرار دهید:

- `pipeline.yml`: قالب Stage مرکزی؛ نمونهٔ آماده در
  `examples/shared-templates/monorepo/pipeline.yml` قرار دارد.
- `mr-build.cjs`: عین فایل `dist/monorepo-build.cjs` این Repository؛ این نسخه
  نام Komodo Stack را نیز در Manifest قرار می‌دهد.
- `nginx/default.conf`: کانفیگ داخلی Runtime استاتیک؛ نمونهٔ آماده در
  `examples/shared-templates/monorepo/nginx/default.conf` قرار دارد. این فایل
  با یک `$uri` معمولی نوشته می‌شود، چون دیگر داخل Compose قرار ندارد.

Pipeline هر پروژه فقط Repositoryهای `SharedTemplatesRepo` و `sourceRepo` را
تعریف می‌کند و پارامترهای Project/Environment را به
`monorepo/pipeline.yml@SharedTemplatesRepo` می‌فرستد. فایل
`/.devops/deployments.yml` همچنان در Repository تولیدشدهٔ همان پروژه باقی
می‌ماند تا قرارداد و استثناهای پروژه مستقل باشند.

Credential مرکزی نمایش Serverها فقط Read است. Credential گروه متغیر
`KomodoAPI` برای این روال باید بتواند Repo و Stack را list/create/update کند،
`DeployStack` را اجرا کند و روی Server انتخاب‌شده Terminal داشته باشد.
