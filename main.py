import os
import csv
import json
import zipfile
from datetime import datetime
import aiosqlite
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    ApplicationBuilder,
    CommandHandler,
    CallbackQueryHandler,
    MessageHandler,
    ContextTypes,
    filters,
)

# ================= ASOSIY SOZLAMALAR =================
BOT_TOKEN = os.environ.get("BOT_TOKEN", "TOKENINGIZNI_SHU_YERGA_QOYING")
SUPER_ADMIN_ID = int(os.environ.get("SUPER_ADMIN_ID", "6887251996"))  # O'zingizning Telegram ID raqamingizni yozing (/myid orqali bilish mumkin)
DB_NAME = "psixolog_bot_v2.db"


# ================= MA'LUMOTLAR BAZASI =================
async def init_db():
    async with aiosqlite.connect(DB_NAME) as db:
        await db.execute("PRAGMA journal_mode=WAL;")

        # Xodimlar jadvali (role: 'admin' yoki 'psychologist')
        await db.execute("""
            CREATE TABLE IF NOT EXISTS staff (
                user_id INTEGER PRIMARY KEY,
                role TEXT,
                name TEXT,
                added_at TEXT
            )
        """)

        # Testlar jadvali
        await db.execute("""
            CREATE TABLE IF NOT EXISTS tests (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                title TEXT,
                group_code TEXT,
                is_anon INTEGER,
                is_auto INTEGER,
                rules_json TEXT DEFAULT '[]',
                created_at TEXT
            )
        """)

        # Savollar jadvali
        await db.execute("""
            CREATE TABLE IF NOT EXISTS questions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_id INTEGER,
                question_text TEXT,
                options_json TEXT
            )
        """)

        # Test natijalari jadvali
        await db.execute("""
            CREATE TABLE IF NOT EXISTS results (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                test_id INTEGER,
                user_id INTEGER,
                student_name TEXT,
                group_code TEXT,
                score INTEGER,
                answers_json TEXT,
                summary TEXT,
                created_at TEXT
            )
        """)

        # Psixologga murojaat va qabul jadvali
        await db.execute("""
            CREATE TABLE IF NOT EXISTS consultations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                student_name TEXT,
                message TEXT,
                reply TEXT DEFAULT '',
                created_at TEXT
            )
        """)
        await db.commit()


# ================= ROLLARNI TEKSHIRISH FUNKSIYALARI =================
async def get_user_role(user_id: int) -> str:
    """Foydalanuvchi rolini qaytaradi: 'superadmin', 'admin', 'psychologist' yoki 'student'"""
    if user_id == SUPER_ADMIN_ID:
        return "superadmin"
    async with aiosqlite.connect(DB_NAME) as db:
        cur = await db.execute("SELECT role FROM staff WHERE user_id = ?", (user_id,))
        row = await cur.fetchone()
        return row[0] if row else "student"


async def get_all_psychologists_and_admins() -> list:
    """Xabarnoma yuborish uchun barcha psixolog va adminlar ID ro'yxati"""
    ids = {SUPER_ADMIN_ID}
    async with aiosqlite.connect(DB_NAME) as db:
        cur = await db.execute("SELECT user_id FROM staff")
        rows = await cur.fetchall()
        for r in rows:
            ids.add(r[0])
    return list(ids)


def cancel_kb():
    """Har bir qadamda orqaga qaytish tugmasi"""
    return InlineKeyboardMarkup([[InlineKeyboardButton("🔙 Bekor qilish / Bosh menyu", callback_data="main_menu")]])


# ================= BOSH MENYU VA /START =================
async def my_id_cmd(update: Update, context: ContextTypes.DEFAULT_TYPE):
    uid = update.effective_user.id
    await update.message.reply_text(
        f"🆔 Sizning Telegram ID raqamingiz:\n`{uid}`\n\n"
        "_(Ushbu raqamni ustiga bosib nusxalab olishingiz va Super Adminga yuborishingiz mumkin)_",
        parse_mode="Markdown"
    )


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data.clear()
    user_id = update.effective_user.id
    role = await get_user_role(user_id)

    # 1. Talabalar uchun tugmalar (hamma ko'radi)
    keyboard = [
        [InlineKeyboardButton("📝 Test topshirish", callback_data="st_start_test")],
        [
            InlineKeyboardButton("📊 Natijalarim", callback_data="st_my_results"),
            InlineKeyboardButton("💬 Psixologga murojaat", callback_data="st_consult")
        ],
    ]

    # 2. Psixolog va Adminlar uchun tugmalar
    if role in ("superadmin", "admin", "psychologist"):
        keyboard.append([InlineKeyboardButton("➕ Yangi test tuzish (Oson rejim)", callback_data="adm_new_test")])
        keyboard.append([InlineKeyboardButton("📥 Shablon olish va Fayl yuklash", callback_data="adm_template_menu")])
        keyboard.append([
            InlineKeyboardButton("📋 Testlar ro'yxati", callback_data="adm_manage_tests"),
            InlineKeyboardButton("📊 Natijalar (Excel)", callback_data="adm_export_excel")
        ])
        keyboard.append([InlineKeyboardButton("📦 ZIP Arxivlash va Bazani tozalash", callback_data="adm_archive_zip")])

    # 3. Faqat Super Admin uchun tugma
    if role == "superadmin":
        keyboard.append([InlineKeyboardButton("👑 Super Admin: Xodimlarni boshqarish", callback_data="sa_staff_menu")])

    role_names = {
        "superadmin": "👑 Super Admin",
        "admin": "🛠 Admin",
        "psychologist": "👨‍⚕️ Psixolog",
        "student": "🎓 Talaba"
    }

    text = (
        f"👋 **Universitet Psixologik Portali**ga xush kelibsiz!\n"
        f"Sizning maqomingiz: **{role_names.get(role, 'Talaba')}**\n\n"
        "Quyidagi bo'limlardan birini tanlang:"
    )

    if update.callback_query:
        await update.callback_query.answer()
        await update.callback_query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")
    else:
        await update.message.reply_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")


# ================= SUPER ADMIN: XODIMLARNI BOSHQARISH =================
async def sa_staff_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    if query.from_user.id != SUPER_ADMIN_ID:
        return

    async with aiosqlite.connect(DB_NAME) as db:
        cur = await db.execute("SELECT user_id, role, name FROM staff")
        staff_list = await cur.fetchall()

    text = "👑 **Xodimlarni boshqarish paneli**\n\n📋 **Hozirgi xodimlar:**\n"
    keyboard = []

    if not staff_list:
        text += "_(Hozircha qo'shimcha admin yoki psixolog yo'q)_\n"
    else:
        for uid, r_name, full_name in staff_list:
            r_icon = "👨‍⚕️ Psixolog" if r_name == "psychologist" else "🛠 Admin"
            text += f"• {r_icon}: **{full_name}** (`{uid}`)\n"
            keyboard.append([InlineKeyboardButton(f"🗑 O'chirish: {full_name} ({r_icon})", callback_data=f"sa_del_{uid}")])

    keyboard.append([
        InlineKeyboardButton("➕ Psixolog qo'shish", callback_data="sa_add_psychologist"),
        InlineKeyboardButton("➕ Admin qo'shish", callback_data="sa_add_admin")
    ])
    keyboard.append([InlineKeyboardButton("🔙 Bosh menyu", callback_data="main_menu")])

    await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")


# ================= SHABLON VA FAYL YUKLASH (TUSHUNARLI INTERFEYS) =================
async def adm_template_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    
    text = (
        "📥 **Fayl orqali test yuklash bo'limi**\n\n"
        "Kompyuterda yoki telefonda savollarni bitta faylga yozib, botga bir zumda yuklashingiz mumkin.\n\n"
        "💡 **Qanday ishlaydi? (3 ta oddiy qadam):**\n"
        "1️⃣ Pastdagi tugmadan **Excel** yoki **CSV** shablonni yuklab oling.\n"
        "2️⃣ Faylni ochib, ichidagi namuna qatorlar o'rniga o'z savollaringizni yozib saqlang.\n"
        "3️⃣ `📤 Tayyor faylni botga yuborish` tugmasini bosing va faylni xuddi oddiy rasm/hujjat yuborgandek chatga tashlang!"
    )
    keyboard = [
        [
            InlineKeyboardButton("📗 Excel (.xlsx) shablon olish", callback_data="adm_get_xlsx"),
            InlineKeyboardButton("📄 CSV (.csv) shablon olish", callback_data="adm_get_csv")
        ],
        [InlineKeyboardButton("📤 Tayyor faylni botga yuborish", callback_data="adm_wait_file")],
        [InlineKeyboardButton("🔙 Bosh menyu", callback_data="main_menu")]
    ]
    await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")


async def adm_send_template_file(update: Update, context: ContextTypes.DEFAULT_TYPE, fmt: str):
    query = update.callback_query
    await query.answer("Shablon tayyorlanmoqda...")

    headers = [
        "Test Nomi",
        "Guruh Kodi (ALL yoki guruh)",
        "Maxfiylik (1=Anonim, 0=Shaxsiy, 2=Ixtiyoriy)",
        "Avto Tahlil (1=Avto, 0=Psixolog tekshiradi)",
        "Savol Matni",
        "Variantlar (Ha=5, Yo'q=0 yoki OPEN)",
        "Xulosa Shabloni (Faqat 1-qatorda)"
    ]
    sample_rows = [
        [
            "Stress darajasi testi", "ALL", 2, 1,
            "Oxirgi haftada uyqungiz qanday?",
            "Yaxshi=0, O'rtacha=3, Yomon=5",
            "0-5: Holatingiz a'lo va barqaror; 6-12: Biroz charchoq bor, dam oling; 13-100: Psixolog maslahatiga keling"
        ],
        [
            "Stress darajasi testi", "ALL", 2, 1,
            "Darslarda diqqatni jamlashga qiynalyapsizmi?",
            "Yo'q=0, Ba'zan=2, Ha=5", ""
        ],
        [
            "Stress darajasi testi", "ALL", 2, 1,
            "Sizni nima ko'proq o'ylantiryapti? (Erkin yozing)",
            "OPEN", ""
        ],
    ]

    if fmt == "xlsx":
        file_path = "Test_Shabloni.xlsx"
        wb = Workbook()
        ws = wb.active
        ws.title = "Shablon"
        ws.append(headers)

        fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        font = Font(color="FFFFFF", bold=True)
        for col_num in range(1, len(headers) + 1):
            c = ws.cell(row=1, column=col_num)
            c.fill = fill
            c.font = font
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        for r in sample_rows:
            ws.append(r)

        widths = [25, 22, 25, 25, 45, 35, 50]
        for i, w in enumerate(widths, 1):
            ws.column_dimensions[chr(64 + i)].width = w
        wb.save(file_path)
    else:
        file_path = "Test_Shabloni.csv"
        with open(file_path, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(headers)
            writer.writerows(sample_rows)

    with open(file_path, "rb") as f:
        await query.message.reply_document(
            document=f,
            caption=(
                f"✅ **Tayyor {fmt.upper()} shablon!**\n\n"
                "📌 **Qanday to'ldiriladi?**\n"
                "• **Guruh Kodi:** Hamma uchun bo'lsa `ALL`, maxsus guruhga bo'lsa `DI-101` kabi yozing.\n"
                "• **Variantlar:** Variant yoniga `=` qo'yib ball yozing (`Ha=5, Yo'q=0`). Agar talaba javobni o'zi yozsa `OPEN` deb yozing.\n\n"
                "To'ldirib bo'lgach, faylni saqlab **xuddi shu chatga qaytarib yuboring**."
            ),
            parse_mode="Markdown"
        )
    os.remove(file_path)


# ================= TESTLARNI BOSHQARISH VA NATIJALARNI OLISH =================
async def adm_manage_tests(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    async with aiosqlite.connect(DB_NAME) as db:
        cur = await db.execute("""
            SELECT t.id, t.title, t.group_code, COUNT(q.id)
            FROM tests t LEFT JOIN questions q ON t.id = q.test_id
            GROUP BY t.id ORDER BY t.id DESC
        """)
        tests = await cur.fetchall()

    if not tests:
        await query.edit_message_text("📭 Hozircha bazada testlar yo'q.", reply_markup=cancel_kb())
        return

    text = "📋 **Mavjud testlar ro'yxati:**\n_(O'chirish uchun test tugmasini bosing)_\n\n"
    keyboard = []
    for tid, title, gcode, qcount in tests:
        text += f"🔹 **{title}** | Guruh: `{gcode}` | Savollar: {qcount} ta\n"
        keyboard.append([InlineKeyboardButton(f"🗑 O'chirish: {title} ({gcode})", callback_data=f"adm_deltest_{tid}")])

    keyboard.append([InlineKeyboardButton("🔙 Bosh menyu", callback_data="main_menu")])
    await query.edit_message_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")


async def adm_export_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    async with aiosqlite.connect(DB_NAME) as db:
        cur = await db.execute("""
            SELECT r.id, t.title, r.student_name, r.group_code, r.score, r.answers_json, r.summary, r.created_at
            FROM results r LEFT JOIN tests t ON r.test_id = t.id
            ORDER BY r.id DESC
        """)
        rows = await cur.fetchall()

    if not rows:
        await query.message.reply_text("📭 Hozircha topshirilgan test natijalari yo'q.")
        return

    file_path = f"Natijalar_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    wb = Workbook()
    ws = wb.active
    ws.title = "Natijalar"
    ws.append(["ID", "Test Nomi", "Talaba", "Guruh", "To'plangan Ball", "Berilgan Javoblar", "Xulosa", "Sana"])

    for r in rows:
        rid, t_title, s_name, g_code, score, ans_json, summ, c_at = r
        try:
            ans_dict = json.loads(ans_json)
            ans_readable = "; ".join([f"{k}: {v}" for k, v in ans_dict.items()])
        except Exception:
            ans_readable = ans_json
        ws.append([rid, t_title or "O'chirilgan test", s_name, g_code, score, ans_readable, summ, c_at])

    wb.save(file_path)
    with open(file_path, "rb") as f:
        await query.message.reply_document(
            document=f,
            caption=f"📊 **Barcha test natijalari (Excel)**\nJami yozuvlar: **{len(rows)}** ta.",
            parse_mode="Markdown"
        )
    os.remove(file_path)


async def adm_archive_zip(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    csv_file = f"results_{timestamp}.csv"
    zip_file = f"Baza_Arxiv_{timestamp}.zip"

    async with aiosqlite.connect(DB_NAME) as db:
        cur = await db.execute("SELECT * FROM results")
        rows = await cur.fetchall()

        if not rows:
            await query.message.reply_text("📭 Bazada arxivlash uchun natijalar yo'q.")
            return

        with open(csv_file, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";")
            writer.writerow(["ID", "Test_ID", "User_ID", "Talaba", "Guruh", "Ball", "Javoblar", "Xulosa", "Sana"])
            writer.writerows(rows)

        with zipfile.ZipFile(zip_file, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.write(csv_file)

        await db.execute("DELETE FROM results")
        await db.commit()
        await db.execute("VACUUM")

    with open(zip_file, "rb") as f:
        await query.message.reply_document(
            document=f,
            caption=f"📦 **Baza ZIP qilib arxivlandi va tozalandi!**\nArxivlangan natijalar: **{len(rows)}** ta.",
            parse_mode="Markdown"
        )
    os.remove(csv_file)
    os.remove(zip_file)


# ================= BARCHA TUGMALAR (CALLBACK) ROUTERI =================
async def callback_router(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    data = query.data
    user_id = query.from_user.id
    role = await get_user_role(user_id)

    # 1. Bosh menyuga qaytish
    if data == "main_menu":
        return await start(update, context)

    # 2. Super Admin funksiyalari
    if data == "sa_staff_menu":
        return await sa_staff_menu(update, context)

    if data in ("sa_add_psychologist", "sa_add_admin") and user_id == SUPER_ADMIN_ID:
        await query.answer()
        target_role = "psychologist" if "psychologist" in data else "admin"
        context.user_data["step"] = "sa_wait_staff_info"
        context.user_data["target_role"] = target_role
        r_title = "Psixolog" if target_role == "psychologist" else "Admin"
        await query.edit_message_text(
            f"➕ **Yangi {r_title} qo'shish**\n\n"
            f"Qo'shmoqchi bo'lgan insonning **Telegram ID raqami** va **Ismini** probel (bo'sh joy) bilan yozib yuboring.\n\n"
            f"📌 **Namuna:**\n`987654321 Dilnoza опа`\n\n"
            f"_(Eslatma: U inson botga kirib `/myid` deb yozsa, bot unga ID raqamini chiqarib beradi)_",
            reply_markup=cancel_kb(),
            parse_mode="Markdown"
        )
        return

    if data.startswith("sa_del_") and user_id == SUPER_ADMIN_ID:
        del_uid = int(data.split("_")[2])
        async with aiosqlite.connect(DB_NAME) as db:
            await db.execute("DELETE FROM staff WHERE user_id = ?", (del_uid,))
            await db.commit()
        await query.answer("Xodim o'chirildi!", show_alert=True)
        return await sa_staff_menu(update, context)

    # 3. Admin & Psixolog funksiyalari
    if data == "adm_template_menu":
        return await adm_template_menu(update, context)
    if data == "adm_get_xlsx":
        return await adm_send_template_file(update, context, "xlsx")
    if data == "adm_get_csv":
        return await adm_send_template_file(update, context, "csv")
    if data == "adm_wait_file":
        await query.answer()
        context.user_data["step"] = "adm_wait_file"
        await query.edit_message_text(
            "📤 **Fayl yuklash rejimi**\n\n"
            "Pastdagi 📎 (skrepka) tugmasini bosib, to'ldirilgan **Excel (`.xlsx`)** yoki **CSV (`.csv`)** faylingizni yuboring:",
            reply_markup=cancel_kb(),
            parse_mode="Markdown"
        )
        return

    if data == "adm_manage_tests":
        return await adm_manage_tests(update, context)

    if data.startswith("adm_deltest_") and role in ("superadmin", "admin", "psychologist"):
        tid = int(data.split("_")[2])
        async with aiosqlite.connect(DB_NAME) as db:
            await db.execute("DELETE FROM tests WHERE id = ?", (tid,))
            await db.execute("DELETE FROM questions WHERE test_id = ?", (tid,))
            await db.commit()
        await query.answer("Test o'chirildi!", show_alert=True)
        return await adm_manage_tests(update, context)

    if data == "adm_export_excel":
        return await adm_export_excel(update, context)
    if data == "adm_archive_zip":
        return await adm_archive_zip(update, context)

    # --- OSON TEST TUZISH BOSQICHLARI ---
    if data == "adm_new_test" and role in ("superadmin", "admin", "psychologist"):
        await query.answer()
        context.user_data["new_test"] = {"questions": []}
        context.user_data["step"] = "nt_title"
        await query.edit_message_text(
            "1️⃣ **Yangi test nomini yozib yuboring:**\n_(Masalan: Stress darajasini aniqlash)_",
            reply_markup=cancel_kb(),
            parse_mode="Markdown"
        )
        return

    if data.startswith("nt_grp_"):
        await query.answer()
        grp_choice = data.split("_")[2]
        if grp_choice == "ALL":
            context.user_data["new_test"]["group_code"] = "ALL"
            return await ask_anon_mode(query.message, context, edit=True)
        else:
            context.user_data["step"] = "nt_custom_group"
            await query.edit_message_text(
                "🎓 **Guruh kodi yoki yo'nalish nomini yozing:**\n_(Masalan: `DI-101` yoki `Dasturiy injiniring`)_",
                reply_markup=cancel_kb(),
                parse_mode="Markdown"
            )
            return

    if data.startswith("nt_anon_"):
        await query.answer()
        context.user_data["new_test"]["is_anon"] = int(data.split("_")[2])
        keyboard = [
            [InlineKeyboardButton("🤖 Avtomatik (Bot o'zi ball hisoblab xulosa beradi)", callback_data="nt_auto_1")],
            [InlineKeyboardButton("👨‍⚕️ Psixolog o'zi o'qib javob yozadi", callback_data="nt_auto_0")],
            [InlineKeyboardButton("🔙 Bekor qilish", callback_data="main_menu")]
        ]
        await query.edit_message_text(
            "4️⃣ **Testni baholash usulini tanlang:**",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )
        return

    if data.startswith("nt_auto_"):
        await query.answer()
        context.user_data["new_test"]["is_auto"] = int(data.split("_")[2])
        context.user_data["step"] = "nt_question_text"
        await query.edit_message_text(
            "5️⃣ **1-savol matnini yozib yuboring:**\n_(Masalan: Kun davomida o'zingizda charchoq sezasizmi?)_",
            reply_markup=cancel_kb(),
            parse_mode="Markdown"
        )
        return

    # Savol variantlarini oson tugmada tanlash
    if data.startswith("qopt_"):
        await query.answer()
        opt_type = data.split("_")[1]
        q_text = context.user_data.get("current_q_text", "Savol")

        if opt_type == "std":
            opts = json.dumps({"Ha": 5, "Ba'zan": 3, "Yo'q": 0}, ensure_ascii=False)
            return await save_question_to_memory(query.message, context, q_text, opts, edit=True)
        elif opt_type == "scale5":
            opts = json.dumps({"1 (Umuman yo'q)": 1, "2": 2, "3 (O'rtacha)": 3, "4": 4, "5 (Juda kuchli)": 5}, ensure_ascii=False)
            return await save_question_to_memory(query.message, context, q_text, opts, edit=True)
        elif opt_type == "open":
            opts = json.dumps("OPEN")
            return await save_question_to_memory(query.message, context, q_text, opts, edit=True)
        elif opt_type == "custom":
            context.user_data["step"] = "nt_custom_options"
            await query.edit_message_text(
                f"❓ **Savol:** _{q_text}_\n\n"
                "Variantlarni vergul bilan ajratib yozing (ball berish uchun `=` qo'ying):\n"
                "📌 **Namuna:** `Doim=5, Ba'zida=2, Hech qachon=0`",
                reply_markup=cancel_kb(),
                parse_mode="Markdown"
            )
            return

    if data == "nt_finish_save":
        await query.answer()
        return await finalize_new_test(query.message, context)

    # --- PSIXOLOG JAVOB YOZISHI (MUROJAAT YOKI TEST NATIJASIGA) ---
    if data.startswith("reply_student_"):
        await query.answer()
        parts = data.split("_")
        target_uid = int(parts[2])
        context.user_data["reply_target_uid"] = target_uid
        context.user_data["step"] = "psych_writing_reply"
        await query.message.reply_text(
            "✍️ **Talabaga yuboriladigan psixologik xulosa yoki javobingizni yozing:**\n"
            "_(Javobingiz talabaga bot orqali darhol yetkaziladi)_",
            reply_markup=cancel_kb(),
            parse_mode="Markdown"
        )
        return

    # --- TALABA: TEST TOPSHIRISH VA MUROJAAT ---
    if data == "st_start_test":
        await query.answer()
        async with aiosqlite.connect(DB_NAME) as db:
            cur = await db.execute("SELECT DISTINCT group_code FROM tests")
            groups = [r[0] for r in await cur.fetchall()]

        if not groups:
            await query.edit_message_text("📭 Hozircha faol testlar mavjud emas.", reply_markup=cancel_kb())
            return

        keyboard = [[InlineKeyboardButton("🌐 Umumiy testlar (Barcha uchun)", callback_data="st_grp_ALL")]]
        for g in groups:
            if g.upper() != "ALL":
                keyboard.append([InlineKeyboardButton(f"🎓 Guruh: {g}", callback_data=f"st_grp_{g}")])
        keyboard.append([InlineKeyboardButton("🔙 Bosh menyu", callback_data="main_menu")])

        await query.edit_message_text(
            "🎓 **O'z guruhingizni yoki Umumiy bo'limni tanlang:**",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )
        return

    if data.startswith("st_grp_"):
        await query.answer()
        g_code = data.split("_", 2)[2]
        context.user_data["student_group"] = g_code

        async with aiosqlite.connect(DB_NAME) as db:
            cur = await db.execute(
                "SELECT id, title, is_anon, is_auto FROM tests WHERE UPPER(group_code) = ? OR UPPER(group_code) = 'ALL'",
                (g_code.upper(),)
            )
            tests = await cur.fetchall()

        if not tests:
            await query.edit_message_text("❌ Bu bo'limda testlar topilmadi.", reply_markup=cancel_kb())
            return

        keyboard = [[InlineKeyboardButton(f"📝 {t[1]}", callback_data=f"st_take_{t[0]}")] for t in tests]
        keyboard.append([InlineKeyboardButton("🔙 Orqaga", callback_data="st_start_test")])
        await query.edit_message_text("📋 **Topshirmoqchi bo'lgan testingizni tanlang:**", reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")
        return

    if data.startswith("st_take_"):
        await query.answer()
        test_id = int(data.split("_")[2])
        async with aiosqlite.connect(DB_NAME) as db:
            cur = await db.execute("SELECT title, is_anon, is_auto, rules_json FROM tests WHERE id = ?", (test_id,))
            t_info = await cur.fetchone()
            cur_q = await db.execute("SELECT question_text, options_json FROM questions WHERE test_id = ?", (test_id,))
            questions = await cur_q.fetchall()

        if not t_info or not questions:
            await query.edit_message_text("⚠️ Bu testda savollar topilmadi.", reply_markup=cancel_kb())
            return

        context.user_data["active_test"] = {
            "test_id": test_id,
            "title": t_info[0],
            "is_anon": t_info[1],
            "is_auto": t_info[2],
            "rules_json": t_info[3],
            "questions": questions,
            "q_idx": 0,
            "score": 0,
            "answers": {},
            "student_name": "Anonim"
        }

        is_anon = t_info[1]
        if is_anon == 1:
            return await send_test_question(query.message, context, edit=True)
        elif is_anon == 0:
            context.user_data["step"] = "st_enter_name"
            await query.edit_message_text("👤 **Ism-familiyangiz va guruhingizni yozib yuboring:**", reply_markup=cancel_kb(), parse_mode="Markdown")
            return
        else:
            keyboard = [
                [InlineKeyboardButton("🕵️ Anonim topshirish (Ismim ko'rinmasin)", callback_data="st_anonchoice_1")],
                [InlineKeyboardButton("👤 Ismimni yozib topshirish", callback_data="st_anonchoice_0")],
                [InlineKeyboardButton("🔙 Bekor qilish", callback_data="main_menu")]
            ]
            await query.edit_message_text(
                "🔒 **Testni qanday shaklda topshirishni xohlaysiz?**",
                reply_markup=InlineKeyboardMarkup(keyboard),
                parse_mode="Markdown"
            )
            return

    if data.startswith("st_anonchoice_"):
        await query.answer()
        choice = int(data.split("_")[2])
        if choice == 1:
            context.user_data["active_test"]["student_name"] = "Anonim"
            return await send_test_question(query.message, context, edit=True)
        else:
            context.user_data["step"] = "st_enter_name"
            await query.edit_message_text("👤 **Ism-familiyangiz va guruhingizni yozib yuboring:**", reply_markup=cancel_kb(), parse_mode="Markdown")
            return

    if data.startswith("st_ans_"):
        await query.answer()
        t_data = context.user_data.get("active_test")
        if not t_data:
            return await start(update, context)

        _, _, opt_idx_str = data.split("_")
        opt_idx = int(opt_idx_str)

        q_text, opts_json = t_data["questions"][t_data["q_idx"]]
        opts_list = list(json.loads(opts_json).items())
        chosen_label, chosen_score = opts_list[opt_idx]

        t_data["score"] += int(chosen_score)
        t_data["answers"][q_text] = chosen_label
        t_data["q_idx"] += 1
        return await send_test_question(query.message, context, edit=True)

    if data == "st_my_results":
        await query.answer()
        async with aiosqlite.connect(DB_NAME) as db:
            cur = await db.execute("""
                SELECT t.title, r.score, r.summary, r.created_at
                FROM results r LEFT JOIN tests t ON r.test_id = t.id
                WHERE r.user_id = ? ORDER BY r.id DESC LIMIT 5
            """, (user_id,))
            rows = await cur.fetchall()

        if not rows:
            await query.edit_message_text("📭 Siz hali test topshirmagansiz.", reply_markup=cancel_kb())
            return

        text = "📊 **Sizning oxirgi 5 ta test natijangiz:**\n\n"
        for title, sc, summ, c_at in rows:
            text += f"📌 **{title or 'Test'}** (`{c_at}`)\n🔹 Ball: **{sc}**\n💡 _{summ}_\n\n"
        await query.edit_message_text(text, reply_markup=cancel_kb(), parse_mode="Markdown")
        return

    if data == "st_consult":
        await query.answer()
        keyboard = [
            [InlineKeyboardButton("🕵️ Anonim murojaat (Shaxsingiz sir qoladi)", callback_data="consult_mode_anon")],
            [InlineKeyboardButton("👤 Ochiq murojaat / Qabulga yozilish", callback_data="consult_mode_open")],
            [InlineKeyboardButton("🔙 Bosh menyu", callback_data="main_menu")]
        ]
        await query.edit_message_text(
            "💬 **Psixologga murojaat va qabulga yozilish**\n\n"
            "Universitet psixologiga o'zingizni qiynayotgan savolni yuborishingiz yoki suhbatga yozilishingiz mumkin. Qanday murojaat qilasiz?",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )
        return

    if data.startswith("consult_mode_"):
        await query.answer()
        mode = data.split("_")[2]
        context.user_data["consult_mode"] = mode
        context.user_data["step"] = "st_writing_consult"
        if mode == "anon":
            prompt_text = "🕵️ **Anonim murojaat:**\nSavolingizni yoki muammongizni batafsil yozib yuboring. Ismingiz psixologga ko'rinmaydi, lekin psixolog javob yozsa bot sizga yetkazadi:"
        else:
            prompt_text = "👤 **Ochiq murojaat / Qabulga yozilish:**\nIsm-familiyangiz, guruhingiz, telefon raqamingiz hamda qachon suhbatga kelmoqchi ekanligingizni yozib yuboring:"
        await query.edit_message_text(prompt_text, reply_markup=cancel_kb(), parse_mode="Markdown")
        return


# ================= MATNLI XABARLAR ROUTERI =================
async def text_router(update: Update, context: ContextTypes.DEFAULT_TYPE):
    step = context.user_data.get("step")
    text = update.message.text.strip()
    user_id = update.effective_user.id

    if not step:
        return await start(update, context)

    # 1. Super Admin xodim qo'shishi
    if step == "sa_wait_staff_info" and user_id == SUPER_ADMIN_ID:
        parts = text.split(maxsplit=1)
        if not parts[0].isdigit():
            await update.message.reply_text("⚠️ Birinchi bo'lib faqat raqamdan iborat Telegram ID yozing!\nNamuna: `123456789 Aziza opa`", reply_markup=cancel_kb(), parse_mode="Markdown")
            return
        new_uid = int(parts[0])
        new_name = parts[1] if len(parts) > 1 else "Xodim"
        t_role = context.user_data.get("target_role", "psychologist")

        async with aiosqlite.connect(DB_NAME) as db:
            await db.execute(
                "INSERT OR REPLACE INTO staff (user_id, role, name, added_at) VALUES (?, ?, ?, ?)",
                (new_uid, t_role, new_name, datetime.now().strftime("%Y-%m-%d %H:%M"))
            )
            await db.commit()

        context.user_data.clear()
        await update.message.reply_text(f"✅ **{new_name}** (`{new_uid}`) muvaffaqiyatli qo'shildi!", parse_mode="Markdown")
        return await start(update, context)

    # 2. Oson test tuzish qadamlari
    if step == "nt_title":
        context.user_data["new_test"]["title"] = text
        context.user_data["step"] = None
        keyboard = [
            [InlineKeyboardButton("🌐 Barcha talabalar uchun (Umumiy)", callback_data="nt_grp_ALL")],
            [InlineKeyboardButton("🎓 Maxsus bitta guruh uchun", callback_data="nt_grp_CUSTOM")],
            [InlineKeyboardButton("🔙 Bekor qilish", callback_data="main_menu")]
        ]
        await update.message.reply_text(
            "2️⃣ **Ushbu test kimlar uchun mo'ljallangan?**",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )
        return

    if step == "nt_custom_group":
        context.user_data["new_test"]["group_code"] = text
        context.user_data["step"] = None
        return await ask_anon_mode(update.message, context, edit=False)

    if step == "nt_question_text":
        context.user_data["current_q_text"] = text
        context.user_data["step"] = None
        keyboard = [
            [InlineKeyboardButton("🔘 Standart (Ha=5 / Ba'zan=3 / Yo'q=0)", callback_data="qopt_std")],
            [InlineKeyboardButton("🔢 5 ballik shkala (1 dan 5 gacha)", callback_data="qopt_scale5")],
            [InlineKeyboardButton("✍️ Ochiq savol (Talaba o'zi fikr yozadi)", callback_data="qopt_open")],
            [InlineKeyboardButton("⚙️ Variantlarni o'zim yozaman", callback_data="qopt_custom")],
            [InlineKeyboardButton("🔙 Bekor qilish", callback_data="main_menu")]
        ]
        await update.message.reply_text(
            f"❓ **Savol:** _{text}_\n\nUshbu savolga qanday javob variantlarini qo'yamiz? Tanlang:",
            reply_markup=InlineKeyboardMarkup(keyboard),
            parse_mode="Markdown"
        )
        return

    if step == "nt_custom_options":
        q_text = context.user_data.get("current_q_text", "Savol")
        opts_json = parse_options(text)
        context.user_data["step"] = None
        return await save_question_to_memory(update.message, context, q_text, opts_json, edit=False)

    # 3. Talaba ism kiritishi va ochiq savolga javob yozishi
    if step == "st_enter_name":
        context.user_data["active_test"]["student_name"] = text
        context.user_data["step"] = None
        return await send_test_question(update.message, context, edit=False)

    if step == "st_open_answer":
        t_data = context.user_data.get("active_test")
        if not t_data:
            return await start(update, context)
        q_text, _ = t_data["questions"][t_data["q_idx"]]
        t_data["answers"][q_text] = text
        t_data["q_idx"] += 1
        context.user_data["step"] = None
        return await send_test_question(update.message, context, edit=False)

    # 4. Talaba psixologga murojaat yozishi
    if step == "st_writing_consult":
        mode = context.user_data.get("consult_mode", "anon")
        s_name = "🕵️ Anonim talaba" if mode == "anon" else f"👤 {update.effective_user.full_name}"

        async with aiosqlite.connect(DB_NAME) as db:
            await db.execute(
                "INSERT INTO consultations (user_id, student_name, message, created_at) VALUES (?, ?, ?, ?)",
                (user_id, s_name, text, datetime.now().strftime("%Y-%m-%d %H:%M"))
            )
            await db.commit()

        # Barcha psixolog va adminlarga xabar yuborish
        staff_ids = await get_all_psychologists_and_admins()
        kb = InlineKeyboardMarkup([[InlineKeyboardButton("✉️ Talabaga javob yozish", callback_data=f"reply_student_{user_id}")]])
        for sid in staff_ids:
            try:
                await context.bot.send_message(
                    chat_id=sid,
                    text=f"📩 **Yangi murojaat keldi!**\n\n**Kimdan:** {s_name}\n**Xabar:** {text}",
                    reply_markup=kb,
                    parse_mode="Markdown"
                )
            except Exception:
                pass

        context.user_data.clear()
        await update.message.reply_text(
            "✅ **Murojaatingiz universitet psixologiga yuborildi!**\nPsixolog javob yozishi bilan bot sizga xabar beradi.",
            parse_mode="Markdown"
        )
        return await start(update, context)

    # 5. Psixolog talabaga javob qaytarishi
    if step == "psych_writing_reply":
        target_uid = context.user_data.get("reply_target_uid")
        try:
            await context.bot.send_message(
                chat_id=target_uid,
                text=f"👨‍⚕️ **Universitet psixologidan javob / tavsiya:**\n\n{text}",
                parse_mode="Markdown"
            )
            await update.message.reply_text("✅ Javobingiz talabaga muvaffaqiyatli yetkazildi!")
        except Exception as e:
            await update.message.reply_text(f"⚠️ Talabaga xabar yuborib bo'lmadi: {e}")

        context.user_data.clear()
        return await start(update, context)


# ================= YORDAMCHI FUNKSIYALAR (TEST VA FAYL) =================
async def ask_anon_mode(msg_obj, context: ContextTypes.DEFAULT_TYPE, edit: bool):
    keyboard = [
        [InlineKeyboardButton("🔄 Talaba o'zi tanlasin (Anonim yoki Ochiq)", callback_data="nt_anon_2")],
        [InlineKeyboardButton("🕵️ Faqat Anonim (Ism so'ralmaydi)", callback_data="nt_anon_1")],
        [InlineKeyboardButton("👤 Majburiy Ochiq (Ism-familiya so'raladi)", callback_data="nt_anon_0")],
        [InlineKeyboardButton("🔙 Bekor qilish", callback_data="main_menu")]
    ]
    text = "3️⃣ **Test maxfiylik darajasini tanlang:**"
    if edit:
        await msg_obj.edit_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")
    else:
        await msg_obj.reply_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")


async def save_question_to_memory(msg_obj, context: ContextTypes.DEFAULT_TYPE, q_text: str, opts_json: str, edit: bool):
    context.user_data["new_test"]["questions"].append((q_text, opts_json))
    count = len(context.user_data["new_test"]["questions"])
    context.user_data["step"] = "nt_question_text"

    keyboard = [
        [InlineKeyboardButton(f"✅ Testni saqlash va yakunlash ({count} ta savol)", callback_data="nt_finish_save")],
        [InlineKeyboardButton("🔙 Bekor qilish", callback_data="main_menu")]
    ]
    text = (
        f"✅ **{count}-savol qo'shildi!**\n\n"
        f"👉 Endi **{count + 1}-savol matnini** yozib yuboring yoki savollar tugagan bo'lsa pastdagi **Saqlash** tugmasini bosing:"
    )
    if edit:
        await msg_obj.edit_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")
    else:
        await msg_obj.reply_text(text, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")


async def finalize_new_test(msg_obj, context: ContextTypes.DEFAULT_TYPE):
    data = context.user_data.get("new_test")
    if not data or not data.get("questions"):
        await msg_obj.edit_text("⚠️ Savollar kiritilmadi.")
        return

    async with aiosqlite.connect(DB_NAME) as db:
        cur = await db.execute(
            "INSERT INTO tests (title, group_code, is_anon, is_auto, rules_json, created_at) VALUES (?, ?, ?, ?, ?, ?)",
            (data["title"], data["group_code"], data["is_anon"], data["is_auto"], "[]", datetime.now().strftime("%Y-%m-%d %H:%M"))
        )
        test_id = cur.lastrowid
        for q_text, opts_json in data["questions"]:
            await db.execute("INSERT INTO questions (test_id, question_text, options_json) VALUES (?, ?, ?)", (test_id, q_text, opts_json))
        await db.commit()

    context.user_data.clear()
    await msg_obj.edit_text(
        f"🎉 **'{data['title']}'** testi muvaffaqiyatli saqlandi!\n"
        f"🎓 Guruh: `{data['group_code']}` | Savollar soni: **{len(data['questions'])}** ta.",
        reply_markup=cancel_kb(),
        parse_mode="Markdown"
    )


async def send_test_question(msg_obj, context: ContextTypes.DEFAULT_TYPE, edit: bool):
    t_data = context.user_data["active_test"]
    q_idx = t_data["q_idx"]
    questions = t_data["questions"]

    if q_idx >= len(questions):
        return await complete_student_test(msg_obj, context, edit)

    q_text, opts_json = questions[q_idx]
    opts = json.loads(opts_json)
    header = f"📝 **{t_data['title']}** ({q_idx + 1}/{len(questions)})\n\n❓ **{q_text}**"

    if opts == "OPEN":
        context.user_data["step"] = "st_open_answer"
        text = f"{header}\n\n✍️ _(Javobingizni yozib yuboring)_"
        if edit:
            await msg_obj.edit_text(text, reply_markup=cancel_kb(), parse_mode="Markdown")
        else:
            await msg_obj.reply_text(text, reply_markup=cancel_kb(), parse_mode="Markdown")
    else:
        keyboard = []
        for idx, (label, _) in enumerate(opts.items()):
            keyboard.append([InlineKeyboardButton(label, callback_data=f"st_ans_{idx}")])
        keyboard.append([InlineKeyboardButton("🔙 Testni to'xtatish", callback_data="main_menu")])

        if edit:
            await msg_obj.edit_text(header, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")
        else:
            await msg_obj.reply_text(header, reply_markup=InlineKeyboardMarkup(keyboard), parse_mode="Markdown")


async def complete_student_test(msg_obj, context: ContextTypes.DEFAULT_TYPE, edit: bool):
    t_data = context.user_data["active_test"]
    score = t_data["score"]
    user_id = msg_obj.chat_id

    # Avtomatik yoki qo'lda tahlil
    summary = ""
    if t_data["is_auto"] == 1:
        rules = json.loads(t_data.get("rules_json") or "[]")
        for r in rules:
            if r["min"] <= score <= r["max"]:
                summary = r["text"]
                break
        if not summary:
            max_possible = len(t_data["questions"]) * 5
            ratio = score / max_possible if max_possible > 0 else 0
            if ratio <= 0.35:
                summary = "Sizning ruhiy holatingiz barqaror va xotirjam. Shu ruhda davom eting!"
            elif ratio <= 0.7:
                summary = "Sizda o'rtacha charchoq yoki yengil zo'riqish bor. Dam olish va uyqu rejimiga e'tibor bering."
            else:
                summary = "Sizda kuchli stress va charchoq alomatlari sezilmoqda. Universitet psixologi bilan suhbatlashishni tavsiya qilamiz."
    else:
        summary = "Javoblaringiz universitet psixologiga yuborildi. Psixolog tahlil qilib, sizga xulosa yuboradi."

    async with aiosqlite.connect(DB_NAME) as db:
        await db.execute(
            "INSERT INTO results (test_id, user_id, student_name, group_code, score, answers_json, summary, created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?)",
            (
                t_data["test_id"], user_id, t_data["student_name"],
                context.user_data.get("student_group", "ALL"), score,
                json.dumps(t_data["answers"], ensure_ascii=False), summary,
                datetime.now().strftime("%Y-%m-%d %H:%M")
            )
        )
        await db.commit()

    # Psixologlarga darhol xabar berish (javob yozish tugmasi bilan)
    staff_ids = await get_all_psychologists_and_admins()
    ans_preview = "\n".join([f"• {k}: *{v}*" for k, v in list(t_data["answers"].items())[:10]])
    kb = InlineKeyboardMarkup([[InlineKeyboardButton("✉️ Talabaga fikr/tahlil yozish", callback_data=f"reply_student_{user_id}")]])
    for sid in staff_ids:
        try:
            await context.bot.send_message(
                chat_id=sid,
                text=(
                    f"📊 **Yangi test topshirildi!**\n"
                    f"📌 Test: **{t_data['title']}**\n"
                    f"👤 Talaba: **{t_data['student_name']}** (`{context.user_data.get('student_group', 'ALL')}`)\n"
                    f"🔢 To'plangan ball: **{score}**\n\n"
                    f"**Javoblar:**\n{ans_preview}"
                ),
                reply_markup=kb,
                parse_mode="Markdown"
            )
        except Exception:
            pass

    result_text = (
        f"✅ **Test yakunlandi!**\n\n"
        f"📌 Test: **{t_data['title']}**\n"
        f"👤 Ism: **{t_data['student_name']}**\n\n"
        f"💡 **Natija va tavsiya:**\n{summary}"
    )
    context.user_data.clear()
    if edit:
        await msg_obj.edit_text(result_text, reply_markup=cancel_kb(), parse_mode="Markdown")
    else:
        await msg_obj.reply_text(result_text, reply_markup=cancel_kb(), parse_mode="Markdown")


# ================= EXCEL VA CSV FAYLLARNI O'QISH =================
def parse_options(opts_raw: str) -> str:
    opts_raw = str(opts_raw).strip()
    if not opts_raw or opts_raw.upper() == "OPEN":
        return json.dumps("OPEN")
    opts_dict = {}
    for item in opts_raw.split(","):
        if "=" in item:
            k, v = item.split("=", 1)
            try:
                opts_dict[k.strip()] = int(v.strip())
            except ValueError:
                opts_dict[k.strip()] = 0
        elif item.strip():
            opts_dict[item.strip()] = 0
    return json.dumps(opts_dict, ensure_ascii=False)


def parse_rules(rules_raw: str) -> str:
    if not rules_raw or not str(rules_raw).strip():
        return json.dumps([])
    rules_list = []
    for part in str(rules_raw).split(";"):
        if ":" in part:
            r_str, txt = part.split(":", 1)
            if "-" in r_str:
                try:
                    mn, mx = [int(x.strip()) for x in r_str.split("-", 1)]
                    rules_list.append({"min": mn, "max": mx, "text": txt.strip()})
                except ValueError:
                    continue
    return json.dumps(rules_list, ensure_ascii=False)


async def handle_file_upload(update: Update, context: ContextTypes.DEFAULT_TYPE):
    role = await get_user_role(update.effective_user.id)
    if role not in ("superadmin", "admin", "psychologist"):
        return

    doc = update.message.document
    fname = doc.file_name.lower()
    if not (fname.endswith(".xlsx") or fname.endswith(".csv")):
        await update.message.reply_text("⚠️ Iltimos, faqat `.xlsx` (Excel) yoki `.csv` formatdagi shablon fayl yuboring.")
        return

    file = await doc.get_file()
    temp_path = f"temp_{doc.file_name}"
    await file.download_to_drive(temp_path)
    rows_data = []

    try:
        if fname.endswith(".xlsx"):
            wb = load_workbook(temp_path, data_only=True)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, values_only=True):
                if row and row[0] and row[4]:
                    rows_data.append(row)
        else:
            with open(temp_path, "r", encoding="utf-8-sig") as f:
                sample = f.read(2048)
                f.seek(0)
                delim = ";" if ";" in sample else ","
                reader = csv.reader(f, delimiter=delim)
                for idx, row in enumerate(reader):
                    if not row or len(row) < 5:
                        continue
                    if idx == 0 and "test" in str(row[0]).lower():
                        continue
                    rows_data.append(row)

        if not rows_data:
            await update.message.reply_text("❌ Fayl ichida savollar topilmadi.")
            return

        created_tests = {}
        total_q = 0
        async with aiosqlite.connect(DB_NAME) as db:
            for row in rows_data:
                title = str(row[0]).strip()
                g_code = str(row[1]).strip() if row[1] else "ALL"
                is_anon = int(row[2]) if row[2] is not None and str(row[2]).isdigit() else 2
                is_auto = int(row[3]) if row[3] is not None and str(row[3]).isdigit() else 1
                q_text = str(row[4]).strip()
                opts_raw = str(row[5]).strip() if len(row) > 5 and row[5] else "OPEN"
                rules_raw = str(row[6]).strip() if len(row) > 6 and row[6] else ""

                if title not in created_tests:
                    cur = await db.execute(
                        "INSERT INTO tests (title, group_code, is_anon, is_auto, rules_json, created_at) VALUES (?, ?, ?, ?, ?, ?)",
                        (title, g_code, is_anon, is_auto, parse_rules(rules_raw), datetime.now().strftime("%Y-%m-%d %H:%M"))
                    )
                    created_tests[title] = cur.lastrowid

                await db.execute(
                    "INSERT INTO questions (test_id, question_text, options_json) VALUES (?, ?, ?)",
                    (created_tests[title], q_text, parse_options(opts_raw))
                )
                total_q += 1
            await db.commit()

        context.user_data.clear()
        await update.message.reply_text(
            f"✅ **Fayldan testlar muvaffaqiyatli yuklandi!**\n"
            f"📌 Testlar soni: **{len(created_tests)}** ta\n"
            f"❓ Jami savollar: **{total_q}** ta",
            reply_markup=cancel_kb(),
            parse_mode="Markdown"
        )
    except Exception as e:
        await update.message.reply_text(f"⚠️ Faylni o'qishda xatolik: `{e}`", parse_mode="Markdown")
    finally:
        if os.path.exists(temp_path):
            os.remove(temp_path)


# ================= ISHGA TUSHIRISH =================
async def post_init(application):
    await init_db()


if __name__ == "__main__":
    app = ApplicationBuilder().token(BOT_TOKEN).post_init(post_init).build()

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("myid", my_id_cmd))
    app.add_handler(CallbackQueryHandler(callback_router))
    app.add_handler(MessageHandler(filters.Document.ALL, handle_file_upload))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_router))

    print("Bot v2.0 ishga tushdi...")
    app.run_polling()
