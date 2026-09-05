const PptxGenJS = require("pptxgenjs");
const path = require("path");

const ROOT = path.join(__dirname, "..");
const OUT = path.join(ROOT, "LIFTX-EMS-App-Guide.pptx");
const FONT = "Tahoma";

const C = {
  bg: "F5F4F1",
  white: "FFFFFF",
  ink: "2B2B29",
  muted: "6F6F69",
  line: "E3E0D9",
  clay: "9C8774",
  claySoft: "EFE8E0",
  bar: "2A2927",
  soft: "FAF9F6",
};

const IMG = {
  qa: path.join(ROOT, "assets", "quick-access-ui.png"),
  eu: path.join(ROOT, "assets", "edge-units-ui.png"),
  bd: path.join(ROOT, "assets", "buildings-ui.png"),
  ls: path.join(ROOT, "assets", "local-servers-ui.png"),
  us: path.join(ROOT, "assets", "users-ui.png"),
};

const pres = new PptxGenJS();
pres.defineLayout({ name: "WIDE", width: 13.333, height: 7.5 });
pres.layout = "WIDE";
pres.author = "LIFTX";
pres.title = "راهنمای کامل EMS App";
pres.subject = "راهنمای داشبورد مدیریتی";
pres.rtlMode = true;

let page = 0;
let TOTAL = 21;

function footer(slide) {
  page += 1;
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 7.22, w: 13.333, h: 0.28, fill: { color: C.bg },
  });
  slide.addText("LIFTX  ·  راهنمای EMS App", {
    x: 0.5, y: 7.22, w: 6, h: 0.26,
    fontFace: FONT, fontSize: 10, color: C.muted, margin: 0,
  });
  slide.addText(String(page).padStart(2, "0"), {
    x: 11.3, y: 7.22, w: 1.5, h: 0.26,
    fontFace: FONT, fontSize: 10, color: C.muted, align: "right", margin: 0,
  });
}

function titleBar(slide, title, subtitle) {
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.333, h: 0.07, fill: { color: C.clay },
  });
  slide.addText(title, {
    x: 0.5, y: 0.22, w: 12.3, h: 0.42,
    fontFace: FONT, fontSize: 26, color: C.ink, bold: true, align: "right", margin: 0,
  });
  if (subtitle) {
    slide.addText(subtitle, {
      x: 0.5, y: 0.64, w: 12.3, h: 0.3,
      fontFace: FONT, fontSize: 13, color: C.muted, align: "right", margin: 0,
    });
  }
}

function darkSlide(title, kicker, line) {
  const slide = pres.addSlide();
  slide.background = { color: C.bar };
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.333, h: 0.07, fill: { color: C.clay },
  });
  if (kicker) {
    slide.addText(kicker, {
      x: 1, y: 2.35, w: 11.3, h: 0.36,
      fontFace: FONT, fontSize: 14, color: C.clay, align: "center", bold: true, charSpacing: 3, margin: 0,
    });
  }
  slide.addText(title, {
    x: 1, y: 2.75, w: 11.3, h: 0.9,
    fontFace: FONT, fontSize: 40, color: C.white, align: "center", bold: true, margin: 0,
  });
  if (line) {
    slide.addText(line, {
      x: 2, y: 3.75, w: 9.3, h: 0.4,
      fontFace: FONT, fontSize: 15, color: "B7B2A9", align: "center", margin: 0,
    });
  }
  footer(slide);
  return slide;
}

function shotSlide(title, subtitle, image) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  titleBar(slide, title, subtitle);
  slide.addShape(pres.ShapeType.roundRect, {
    x: 0.45, y: 1.08, w: 12.4, h: 5.95,
    fill: { color: C.white },
    line: { color: C.line, width: 1 },
    rectRadius: 0.1,
  });
  slide.addImage({
    path: image,
    x: 0.6, y: 1.22, w: 12.1, h: 5.67,
    sizing: { type: "contain", w: 12.1, h: 5.67 },
  });
  footer(slide);
}

function cardsSlide(title, subtitle, items) {
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  titleBar(slide, title, subtitle);

  const cols = 2;
  const rows = Math.ceil(items.length / cols);
  const gapX = 0.22;
  const gapY = 0.16;
  const top = 1.1;
  const left = 0.45;
  const areaW = 12.4;
  const areaH = 5.9;
  const cardW = (areaW - gapX) / cols;
  const cardH = (areaH - gapY * (rows - 1)) / rows;

  items.forEach((item, i) => {
    const col = i % cols;
    const row = Math.floor(i / cols);
    const x = left + (cols - 1 - col) * (cardW + gapX);
    const y = top + row * (cardH + gapY);
    slide.addShape(pres.ShapeType.roundRect, {
      x, y, w: cardW, h: cardH,
      fill: { color: C.white },
      line: { color: C.line, width: 1 },
      rectRadius: 0.1,
    });
    slide.addShape(pres.ShapeType.rect, {
      x: x + cardW - 0.08, y: y + 0.18, w: 0.07, h: Math.min(0.42, cardH - 0.36),
      fill: { color: C.clay },
    });
    slide.addText(item[0], {
      x: x + 0.22, y: y + 0.14, w: cardW - 0.46, h: 0.34,
      fontFace: FONT, fontSize: 15, color: C.ink, bold: true, align: "right", margin: 0,
    });
    slide.addText(item[1], {
      x: x + 0.22, y: y + 0.48, w: cardW - 0.46, h: cardH - 0.64,
      fontFace: FONT, fontSize: 13, color: C.muted, align: "right", margin: 0, valign: "top",
    });
  });
  footer(slide);
}

// 1 Cover
{
  const slide = pres.addSlide();
  slide.background = { color: C.bar };
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.333, h: 0.07, fill: { color: C.clay },
  });
  slide.addText("LIFTX", {
    x: 1, y: 1.7, w: 11.3, h: 0.4,
    fontFace: FONT, fontSize: 16, color: C.clay, align: "center", bold: true, charSpacing: 6, margin: 0,
  });
  slide.addText("EMS App", {
    x: 1, y: 2.15, w: 11.3, h: 1.05,
    fontFace: FONT, fontSize: 58, color: C.white, align: "center", bold: true, margin: 0,
  });
  slide.addText("راهنمای کامل سامانه", {
    x: 1, y: 3.25, w: 11.3, h: 0.48,
    fontFace: FONT, fontSize: 24, color: "D5D1C8", align: "center", margin: 0,
  });
  slide.addText("معرفی بخش‌های داشبورد مدیریتی و نحوه کار با هر صفحه", {
    x: 2.2, y: 3.85, w: 8.9, h: 0.36,
    fontFace: FONT, fontSize: 14, color: "9C978E", align: "center", margin: 0,
  });
  slide.addShape(pres.ShapeType.rect, {
    x: 6.05, y: 4.4, w: 1.25, h: 0.03, fill: { color: C.clay },
  });
  slide.addText("دسترسی سریع     واحدهای لبه     ساختمان‌ها     سرورهای محلی     کاربران", {
    x: 1, y: 4.7, w: 11.3, h: 0.36,
    fontFace: FONT, fontSize: 13, color: "B7B2A9", align: "center", margin: 0,
  });
  footer(slide);
}

// 2 Contents
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  titleBar(slide, "فهرست راهنما", "مسیر مطالعه این فایل از ابتدا تا انتها");
  const toc = [
    ["معرفی داشبورد", "منوی اصلی و ارتباط بخش‌ها"],
    ["دسترسی سریع", "واحدهای پین‌شده و کارت هر واحد"],
    ["واحدهای لبه", "ایجاد، جستجو، فیلتر و وضعیت"],
    ["ساختمان‌ها", "کد، آدرس، طبقات و واحدهای متصل"],
    ["سرورهای محلی", "سرور فعال و سرور جایگزین"],
    ["کاربران", "حساب، نقش و اطلاعات تماس"],
  ];
  toc.forEach((item, i) => {
    const y = 1.15 + i * 0.9;
    slide.addShape(pres.ShapeType.roundRect, {
      x: 1.4, y, w: 10.5, h: 0.78,
      fill: { color: C.white },
      line: { color: C.line, width: 1 },
      rectRadius: 0.08,
    });
    slide.addText(item[0], {
      x: 2.1, y: y + 0.1, w: 9.4, h: 0.32,
      fontFace: FONT, fontSize: 18, color: C.ink, bold: true, align: "right", margin: 0,
    });
    slide.addText(item[1], {
      x: 2.1, y: y + 0.4, w: 9.4, h: 0.26,
      fontFace: FONT, fontSize: 13, color: C.muted, align: "right", margin: 0,
    });
  });
  footer(slide);
}

// 3 Dashboard intro
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  titleBar(slide, "داشبورد مدیریتی", "منوی سمت راست در تمام صفحات یکسان است");
  const menus = [
    ["دسترسی سریع", "ورود سریع به واحدهایی که پین شده‌اند"],
    ["واحدهای لبه", "مدیریت کامل دستگاه‌ها و آسانسورهای هر واحد"],
    ["ساختمان‌ها", "ثبت و مشاهده ساختمان‌ها و مشخصات آن‌ها"],
    ["سرورهای محلی", "سرور هر ساختمان و واحدهای متصل به آن"],
    ["کاربران", "حساب‌ها، نقش‌ها و اطلاعات تماس"],
    ["تماس با ما", "ارتباط با تیم پشتیبانی"],
  ];
  menus.forEach((item, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + (1 - col) * 6.4;
    const y = 1.2 + row * 1.85;
    slide.addShape(pres.ShapeType.roundRect, {
      x, y, w: 6.15, h: 1.68,
      fill: { color: C.white },
      line: { color: C.line, width: 1 },
      rectRadius: 0.1,
    });
    slide.addText(item[0], {
      x: x + 0.28, y: y + 0.35, w: 5.6, h: 0.4,
      fontFace: FONT, fontSize: 20, color: C.ink, bold: true, align: "right", margin: 0,
    });
    slide.addText(item[1], {
      x: x + 0.28, y: y + 0.85, w: 5.6, h: 0.5,
      fontFace: FONT, fontSize: 14, color: C.muted, align: "right", margin: 0,
    });
  });
  footer(slide);
}

// 4 How parts connect
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  titleBar(slide, "ارتباط بخش‌ها", "هر ساختمان سرور محلی و واحد لبه دارد؛ کاربر با نقش مشخص وارد می‌شود");
  const steps = [
    ["ساختمان", "محل فیزیکی آسانسورها و آدرس ثبت‌شده"],
    ["سرور محلی", "سرور همان ساختمان برای ارتباط واحدها"],
    ["واحد لبه", "دستگاه هر گروه و تعداد آسانسورها"],
    ["دسترسی سریع", "پین کردن واحدهای پرکاربرد"],
    ["کاربر", "ورود با نقش Admin، Manager یا Maintainer"],
  ];
  steps.forEach((item, i) => {
    const x = 0.4 + i * 2.58;
    slide.addShape(pres.ShapeType.roundRect, {
      x, y: 2.15, w: 2.42, h: 3.4,
      fill: { color: C.white },
      line: { color: C.line, width: 1 },
      rectRadius: 0.1,
    });
    slide.addShape(pres.ShapeType.ellipse, {
      x: x + 0.86, y: 2.45, w: 0.7, h: 0.7,
      fill: { color: C.claySoft },
    });
    slide.addText(String(i + 1), {
      x: x + 0.86, y: 2.45, w: 0.7, h: 0.7,
      fontFace: FONT, fontSize: 18, color: C.ink, align: "center", valign: "middle", bold: true, margin: 0,
    });
    slide.addText(item[0], {
      x: x + 0.14, y: 3.35, w: 2.14, h: 0.7,
      fontFace: FONT, fontSize: 16, color: C.ink, align: "center", bold: true, margin: 0,
    });
    slide.addText(item[1], {
      x: x + 0.14, y: 4.1, w: 2.14, h: 1.15,
      fontFace: FONT, fontSize: 13, color: C.muted, align: "center", margin: 0,
    });
  });
  footer(slide);
}

darkSlide("دسترسی سریع", "بخش اول", "واحدهای پین‌شده برای ورود فوری");
shotSlide("نمای صفحه دسترسی سریع", "واحدها بر اساس ساختمان گروه‌بندی شده‌اند", IMG.qa);
cardsSlide("راهنمای دسترسی سریع", "هر کارت واحد، وضعیت و تعداد آسانسور را نشان می‌دهد", [
  ["ایجاد جدید", "واحد لبه تازه را از همین صفحه تعریف کنید تا در سامانه ثبت شود و بعداً در فهرست کامل هم دیده شود."],
  ["گروه‌های واحد", "واحدها زیر نام ساختمان قرار می‌گیرند؛ مثلاً هدیش، مهرگان یا پالادیوم تا پیدا کردنشان ساده‌تر باشد."],
  ["کارت واحد", "روی کارت، کد واحد مثل E102، عنوان گروه، وضعیت آنلاین و تعداد آسانسور همان واحد دیده می‌شود."],
  ["وضعیت آنلاین", "نقطه سبز کنار کد یعنی واحد الان متصل و در دسترس است. اگر سبز نباشد، واحد آفلاین است."],
  ["تعداد آسانسورها", "عدد پایین کارت نشان می‌دهد چند آسانسور به آن واحد وصل شده‌اند."],
  ["پین دسترسی سریع", "واحدهای پرکاربرد را پین کنید تا همیشه در این صفحه باشند و لازم نباشد هر بار در فهرست بگردید."],
]);

darkSlide("واحدهای لبه", "بخش دوم", "فهرست کامل، فیلتر و مدیریت وضعیت");
shotSlide("نمای صفحه واحدهای لبه", "جدول همه واحدها با جستجو، فیلتر و صفحه‌بندی", IMG.eu);
cardsSlide("راهنمای واحدهای لبه", "از این جدول واحد بسازید، پیدا کنید و وضعیتش را عوض کنید", [
  ["ایجاد واحد جدید", "با دکمه ایجاد جدید یک واحد لبه تازه می‌سازید و ساختمان، عنوان و توضیحات را برایش ثبت می‌کنید."],
  ["جستجو و فیلتر", "کادر فیلتر بر اساس عنوان کمک می‌کند بین واحدهای زیاد سریع همان مورد را پیدا کنید."],
  ["نمایش واحدهای غیرفعال", "اگر این سوییچ روشن باشد، واحدهایی که غیرفعال شده‌اند هم در جدول دیده می‌شوند."],
  ["فقط آنلاین", "با این سوییچ فقط واحدهای متصل نمایش داده می‌شوند تا وضعیت شبکه را سریع‌تر ببینید."],
  ["تعداد ردیف و صفحه‌بندی", "تعداد ردیف هر صفحه را عوض کنید و با فلش‌ها بین صفحات فهرست جابه‌جا شوید."],
  ["ستون‌های جدول", "کد، ساختمان، عنوان، توضیحات، تاریخ ایجاد، فعال بودن و پین دسترسی سریع در هر ردیف آمده است."],
  ["وضعیت آنلاین", "آیکون مانیتور در ابتدای ردیف مشخص می‌کند واحد الان آنلاین است یا نه."],
  ["فعال و دسترسی سریع", "از همین جدول می‌توانید واحد را فعال یا غیرفعال کنید و به دسترسی سریع اضافه کنید."],
]);

darkSlide("ساختمان‌ها", "بخش سوم", "مشخصات محل و واحدهای متصل");
shotSlide("نمای صفحه ساختمان‌ها", "کد، نام، آدرس، طبقات و واحدهای لبه هر ساختمان", IMG.bd);
cardsSlide("راهنمای ساختمان‌ها", "هر ساختمان یک کد یکتا و فهرست واحدهای متصل دارد", [
  ["ایجاد ساختمان جدید", "ساختمان تازه را با نام، کد، آدرس و تعداد طبقات ثبت کنید تا بعداً واحد و سرور به آن وصل شود."],
  ["نام ساختمان", "نام رسمی ساختمان در سامانه؛ معمولاً همان چیزی است که در گزارش‌ها و گروه‌بندی واحدها دیده می‌شود."],
  ["کد ساختمان", "کد کوتاه و یکتا مثل RNKPLC یا PLSHTL برای شناسایی سریع ساختمان در جدول‌ها و اتصال‌ها."],
  ["آدرس", "نشانی ثبت‌شده ساختمان، به فارسی یا انگلیسی، برای پیدا کردن محل روی نقشه یا در پشتیبانی."],
  ["طبقات", "تعداد طبقات ثبت‌شده همان ساختمان؛ برای درک مقیاس پروژه و ظرفیت آسانسور مفید است."],
  ["واحدهای لبه", "برچسب‌هایی مثل E14 یا E19 نشان می‌دهند کدام واحدها به این ساختمان وصل شده‌اند."],
  ["تاریخ ایجاد", "زمان ثبت ساختمان در سامانه تا سابقه تغییرات مشخص بماند."],
  ["کاربرد این صفحه", "قبل از ساخت واحد لبه یا سرور محلی، ساختمان باید اینجا تعریف شده باشد."],
]);

darkSlide("سرورهای محلی", "بخش چهارم", "ارتباط ساختمان با واحدهای لبه");
shotSlide("نمای صفحه سرورهای محلی", "عنوان سرور، ساختمان، واحدهای متصل و وضعیت فعال", IMG.ls);
cardsSlide("راهنمای سرورهای محلی", "هر سرور به یک ساختمان و چند واحد لبه وصل می‌شود", [
  ["ایجاد سرور جدید", "سرور محلی تازه را بسازید و ساختمان مربوط را برایش انتخاب کنید."],
  ["عنوان سرور", "نام خوانا مثل Palladium North Local Server تا تیم فنی سریع بفهمد این سرور مال کدام بخش است."],
  ["ساختمان", "مشخص می‌کند سرور مال کدام ساختمان است. اگر ساختمان نداشته باشد، معمولاً N/A دیده می‌شود."],
  ["واحدهای لبه", "کد واحدهایی که به این سرور وصل شده‌اند روی ردیف با برچسب نمایش داده می‌شود."],
  ["وضعیت سرور", "نقطه سبز یعنی سرور فعال و در دسترس است. بدون آن، ارتباط واحدها ممکن است قطع باشد."],
  ["سرور جایگزین", "Place Holder برای واحدهایی است که هنوز سرور مشخصی ندارند؛ موقتی است تا سرور اصلی تعریف شود."],
]);

darkSlide("کاربران", "بخش پنجم", "حساب ورود، نقش و اطلاعات تماس");
shotSlide("نمای صفحه کاربران", "فهرست حساب‌ها با نام کاربری، تماس، تاریخ و نقش", IMG.us);
cardsSlide("راهنمای کاربران", "هر فرد با یک نقش مشخص وارد سامانه می‌شود", [
  ["ایجاد کاربر جدید", "حساب تازه بسازید و نام کاربری، مشخصات، تماس و نقش را برایش ثبت کنید."],
  ["نام کاربری", "همان نامی است که کاربر هنگام ورود به سامانه وارد می‌کند."],
  ["نام و نام خانوادگی", "مشخصات فردی ثبت‌شده برای شناسایی صاحب حساب در فهرست و پشتیبانی."],
  ["شماره تلفن و ایمیل", "راه‌های تماس ثبت‌شده برای اطلاع‌رسانی یا بازیابی دسترسی."],
  ["تاریخ ایجاد", "زمان ساخته شدن حساب تا مشخص باشد از کی در سامانه فعال شده است."],
  ["نقش", "سطح دسترسی مثل Admin، Manager، Maintainer یا کاربر عادی را تعیین می‌کند."],
  ["صفحه‌بندی", "اگر تعداد کاربران زیاد باشد، با صفحه‌بندی بین صفحات فهرست جابه‌جا شوید."],
  ["فیلتر نام کاربری", "با کادر فیلتر می‌توانید سریع یک حساب را بین همه کاربران پیدا کنید."],
]);

// Roles
{
  const slide = pres.addSlide();
  slide.background = { color: C.bg };
  titleBar(slide, "نقش‌های کاربری", "سطح دسترسی مشخص می‌کند هر نفر چه کاری در سامانه می‌تواند بکند");
  const roles = [
    ["Admin", "دسترسی کامل به تنظیمات، کاربران و همه بخش‌های داشبورد"],
    ["Manager", "مدیریت ساختمان‌ها، واحدها و سرورها در محدوده مسئولیتی"],
    ["Maintainer", "پیگیری وضعیت واحدها، آنلاین بودن و موارد نگهداری"],
    ["کاربر", "مشاهده بخش‌های مجاز بدون تغییر تنظیمات اصلی"],
  ];
  roles.forEach((item, i) => {
    const y = 1.2 + i * 1.4;
    slide.addShape(pres.ShapeType.roundRect, {
      x: 1.15, y, w: 11, h: 1.22,
      fill: { color: C.white },
      line: { color: C.line, width: 1 },
      rectRadius: 0.1,
    });
    slide.addText(item[0], {
      x: 1.5, y: y + 0.22, w: 10.3, h: 0.36,
      fontFace: FONT, fontSize: 20, color: C.ink, bold: true, align: "right", margin: 0,
    });
    slide.addText(item[1], {
      x: 1.5, y: y + 0.64, w: 10.3, h: 0.36,
      fontFace: FONT, fontSize: 14, color: C.muted, align: "right", margin: 0,
    });
  });
  footer(slide);
}

// End
{
  const slide = pres.addSlide();
  slide.background = { color: C.bar };
  slide.addShape(pres.ShapeType.rect, {
    x: 0, y: 0, w: 13.333, h: 0.07, fill: { color: C.clay },
  });
  slide.addText("EMS App  ·  LIFTX", {
    x: 1, y: 2.05, w: 11.3, h: 0.36,
    fontFace: FONT, fontSize: 15, color: C.clay, align: "center", bold: true, margin: 0,
  });
  slide.addText("پایان راهنما", {
    x: 1, y: 2.5, w: 11.3, h: 0.8,
    fontFace: FONT, fontSize: 44, color: C.white, align: "center", bold: true, margin: 0,
  });
  slide.addText("برای پشتیبانی و راهنمایی بیشتر با تیم تماس بگیرید", {
    x: 2, y: 3.4, w: 9.3, h: 0.4,
    fontFace: FONT, fontSize: 16, color: "B7B2A9", align: "center", margin: 0,
  });
  slide.addShape(pres.ShapeType.roundRect, {
    x: 4.45, y: 4.1, w: 4.4, h: 0.72,
    fill: { color: "3B3A36" },
    rectRadius: 0.12,
  });
  slide.addText("09107727044", {
    x: 4.45, y: 4.1, w: 4.4, h: 0.72,
    fontFace: FONT, fontSize: 22, color: C.white, align: "center", valign: "middle", bold: true, margin: 0,
  });
  footer(slide);
}

TOTAL = page;

pres.writeFile({ fileName: OUT }).then(() => {
  console.log("WROTE", OUT, "PAGES", page);
}).catch((error) => {
  console.error(error);
  process.exit(1);
});
