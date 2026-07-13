// User-manual generator for StudentSIS Lao.
// Produces a Word .docx and (if Word/LibreOffice is installed) auto-converts to
// PDF. Word is used as the renderer because iTextSharp 5 doesn't support the
// OpenType shaping needed to position Lao combining vowel marks correctly —
// Word handles complex scripts natively, so the .docx route gives perfect Lao
// rendering.
//
// Invocation:
//   dotnet run -- manual <output-path>
// If <output-path> ends in .pdf, we generate a temp .docx and convert.
// If it ends in .docx (or is omitted), we just produce the .docx and document
// how to make a PDF from it.
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace StudentSIS.IntegrationTests;

internal static class UserManual
{
    private const string LaoFontName = "Phetsarath OT"; // primary; Word falls back if missing
    private const string LaoFontAlt  = "Saysettha OT";
    private const string LaoFontEng  = "Leelawadee UI";

    public static void Generate(string outPath)
    {
        bool wantPdf = outPath.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase);
        string docxPath = wantPdf
            ? Path.Combine(Path.GetTempPath(), $"sis_manual_{Guid.NewGuid():N}.docx")
            : outPath;

        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(outPath))!);
        BuildDocx(docxPath);
        Console.WriteLine($"  Word doc: {docxPath}  ({new FileInfo(docxPath).Length:N0} bytes)");

        if (!wantPdf)
        {
            Console.WriteLine($"WROTE: {outPath}");
            Console.WriteLine();
            Console.WriteLine("  To make a PDF from this:");
            Console.WriteLine("    1. Open the .docx in Microsoft Word");
            Console.WriteLine("    2. File → Save As → choose 'PDF (*.pdf)'");
            Console.WriteLine("    OR use the auto-converter:");
            Console.WriteLine($"    dotnet run --project tests/IntegrationTests -- manual installer/User-Manual.pdf");
            return;
        }

        Console.WriteLine("  Converting .docx → .pdf via Word COM...");
        if (TryConvertViaWord(docxPath, outPath, out string err))
        {
            try { File.Delete(docxPath); } catch { }
            Console.WriteLine($"WROTE: {outPath}  ({new FileInfo(outPath).Length:N0} bytes)");
            return;
        }
        Console.WriteLine($"  Word COM unavailable: {err}");
        Console.WriteLine("  Trying LibreOffice headless...");
        if (TryConvertViaLibreOffice(docxPath, outPath, out err))
        {
            try { File.Delete(docxPath); } catch { }
            Console.WriteLine($"WROTE: {outPath}  ({new FileInfo(outPath).Length:N0} bytes)");
            return;
        }
        Console.WriteLine($"  LibreOffice unavailable: {err}");
        // Last-ditch: keep the .docx alongside the requested .pdf path so user can convert manually.
        string fallback = Path.ChangeExtension(outPath, ".docx");
        File.Copy(docxPath, fallback, true);
        try { File.Delete(docxPath); } catch { }
        Console.WriteLine();
        Console.WriteLine($"  Could not produce PDF automatically. Word document saved instead:");
        Console.WriteLine($"    {fallback}");
        Console.WriteLine($"  Open it in Word and File → Save As → PDF.");
    }

    // ─────────────────────────────────────────────────────────────
    //  .docx builder
    // ─────────────────────────────────────────────────────────────
    private static void BuildDocx(string path)
    {
        if (File.Exists(path)) File.Delete(path);
        using var doc = WordprocessingDocument.Create(path, WordprocessingDocumentType.Document);
        var mainPart = doc.AddMainDocumentPart();
        mainPart.Document = new Document(new Body());
        var body = mainPart.Document.Body!;

        // ── Cover page ────────────────────────────────────────
        for (int i = 0; i < 5; i++) body.AppendChild(BlankLine());
        body.AppendChild(P("ຄູ່ມືການໃຊ້ງານ",  size: 48, bold: true, center: true, color: "1B4F8A"));
        body.AppendChild(P("ລະບົບຄຸ້ມຄອງຂໍ້ມູນນັກຮຽນ", size: 36, bold: true, center: true, color: "1B4F8A"));
        body.AppendChild(BlankLine());
        body.AppendChild(P("StudentSIS Lao  ·  ເວີຊັນ 1.0.0", size: 28, center: true, color: "4B5563"));
        body.AppendChild(BlankLine());
        body.AppendChild(P("ໂຮງຮຽນ ມສ ບຶກທົ່ງ (ມ.1 – ມ.4)", size: 28, center: true, color: "4B5563"));
        for (int i = 0; i < 14; i++) body.AppendChild(BlankLine());
        body.AppendChild(P($"ສ້າງເມື່ອ: {DateTime.Now:dd/MM/yyyy}",
            size: 20, italic: true, center: true, color: "9CA3AF"));
        body.AppendChild(PageBreak());

        // ── Table of contents (no page numbers — Word will compute on Save-As-PDF) ──
        H1(body, "ສາລະບານ");
        var toc = new[] {
            "1.  ພາບລວມຂອງລະບົບ",
            "2.  ການເຂົ້າສູ່ລະບົບ",
            "3.  ໜ້າຫຼັກ (Dashboard)",
            "4.  ການຈັດການນັກຮຽນ",
            "5.  ການລົງທະບຽນວິຊາ",
            "6.  ການບັນທຶກຄະແນນປະຈຳເດືອນ",
            "7.  ການບັນທຶກຄະແນນປະຈຳພາກ",
            "8.  ຄຸນສົມບັດ ແລະ ການອອກແຮງງານ",
            "9.  ການອອກລາຍງານ",
            "10. ການຂຶ້ນຊັ້ນ ແລະ ການຈົບການສຶກສາ",
            "11. ການຈັດການປີການສຶກສາ",
            "12. ການຈັດການວິຊາ ແລະ ຜູ້ໃຊ້ລະບົບ",
            "13. ການຕັ້ງຄ່າລະບົບ / Backup ແລະ Restore",
            "14. ການແກ້ໄຂບັນຫາທີ່ພົບເລື້ອຍ",
        };
        foreach (var line in toc) body.AppendChild(P("   " + line, size: 22));
        body.AppendChild(PageBreak());

        // ── 1. Overview ───────────────────────────────────────
        H1(body, "1. ພາບລວມຂອງລະບົບ");
        Body(body, "ລະບົບ StudentSIS Lao ເປັນລະບົບຄຸ້ມຄອງຂໍ້ມູນນັກຮຽນແບບ offline ສຳລັບ ໂຮງຮຽນມັດທະຍົມສຶກສາຕອນຕົ້ນ ບຶກທົ່ງ (ມ.1 – ມ.4). ໃຊ້ໄດ້ໂດຍບໍ່ຕ້ອງເຊື່ອມຕໍ່ອິນເຕີເນັດ. ຂໍ້ມູນທັງໝົດເກັບໄວ້ໃນໄຟລ໌ດຽວຊື່ \"sis_lao.db\" ທີ່ຢູ່ໃນໂຟນເດີຕິດຕັ້ງ.");
        H2(body, "ໜ້າທີ່ຫຼັກຂອງລະບົບ:");
        Bullet(body, "ບັນທຶກ ແລະ ຈັດການຂໍ້ມູນສ່ວນຕົວນັກຮຽນ (ເພີ່ມ / ແກ້ໄຂ / ລົບ)");
        Bullet(body, "ລົງທະບຽນວິຊາ (ລາຍຄົນ ຫຼື ທັງຫ້ອງ)");
        Bullet(body, "ບັນທຶກຄະແນນປະຈຳເດືອນ (4 ເດືອນຕໍ່ພາກ) ແລະ ຄະແນນປະຈຳພາກ");
        Bullet(body, "ບັນທຶກຄຸນສົມບັດ ແລະ ການອອກແຮງງານ (ປ້ອນດ້ວຍຕົນເອງ)");
        Bullet(body, "ອອກລາຍງານ Excel ແລະ PDF: ປະຈຳເດືອນ, ສະຫຼຸບພາກ, ປະຈຳປີ, ໃບຄະແນນບຸກຄົນ, ໃບສັນຍາ, ລາຍງານປະຫວັດ");
        Bullet(body, "ຂຶ້ນຊັ້ນ / ຈົບການສຶກສາ ແບບຫຼາຍຄົນພ້ອມກັນ");
        Bullet(body, "Backup ແລະ Restore ຖານຂໍ້ມູນ");

        H2(body, "ບົດບາດຜູ້ໃຊ້:");
        Body(body, "ມີ 2 ບົດບາດ:");
        Bullet(body, "ຜູ້ດູແລລະບົບ (admin): ໃຊ້ໄດ້ທຸກໜ້າ, ລວມທັງ ຈັດການຜູ້ໃຊ້, ຂຶ້ນຊັ້ນ, ປີການສຶກສາ, ຕັ້ງຄ່າລະບົບ");
        Bullet(body, "ຄູ (teacher): ໃຊ້ໄດ້ສະເພາະໜ້າທີ່ກ່ຽວຂ້ອງກັບການບັນທຶກຄະແນນ ແລະ ການອອກລາຍງານ");
        Tip(body, "ບັນຊີຜູ້ໃຊ້ເລີ່ມຕົ້ນ: admin / admin1234 — ກະລຸນາປ່ຽນລະຫັດຜ່ານທັນທີຫຼັງຕິດຕັ້ງຄັ້ງທຳອິດ.");
        body.AppendChild(PageBreak());

        // ── 2. Login ──────────────────────────────────────────
        H1(body, "2. ການເຂົ້າສູ່ລະບົບ");
        Body(body, "ເມື່ອເປີດໂປຣແກຣມ ໜ້າ Login ຈະປະກົດຂຶ້ນ. ປ້ອນຊື່ຜູ້ໃຊ້ ແລະ ລະຫັດຜ່ານ ແລ້ວກົດ \"ເຂົ້າສູ່ລະບົບ\". ຖ້າຂໍ້ມູນຖືກຕ້ອງ ໜ້າຫຼັກຈະປະກົດຂຶ້ນພ້ອມເມນູດ້ານຊ້າຍ.");
        H2(body, "ຂັ້ນຕອນ:");
        Step(body, "1.", "ດັບເບີນຄຣິກ (double-click) ໄອຄອນ \"StudentSIS Lao\" ເທິງ Desktop ຫຼື ໃນ Start Menu.");
        Step(body, "2.", "ໃສ່ \"ຊື່ຜູ້ໃຊ້\" (ຕົວຢ່າງ: admin).");
        Step(body, "3.", "ໃສ່ \"ລະຫັດຜ່ານ\".");
        Step(body, "4.", "ກົດປຸ່ມ \"ເຂົ້າສູ່ລະບົບ\".");
        Tip(body, "ຖ້າລະຫັດຜ່ານຜິດ ຈະປະກົດຂໍ້ຄວາມເຕືອນ — ກວດສອບໃໝ່ ຫຼື ຕິດຕໍ່ admin.");
        body.AppendChild(PageBreak());

        // ── 3. Dashboard ──────────────────────────────────────
        H1(body, "3. ໜ້າຫຼັກ (Dashboard)");
        Body(body, "ໜ້າຫຼັກສະແດງສະຖິຕິສຳຄັນ ແລະ ປ່ອງສະຫຼຸບໃຫ້ຮູ້ສະພາບລວມຂອງໂຮງຮຽນໃນປະຈຸບັນ.");
        H2(body, "ສິ່ງທີ່ສະແດງ:");
        Bullet(body, "ນັກຮຽນທັງໝົດ — ນັບທຸກສະຖານະ");
        Bullet(body, "ນັກຮຽນກຳລັງຮຽນ — ສະຖານະ ‘ກຳລັງຮຽນ’ ເທົ່ານັ້ນ");
        Bullet(body, "ຈຳນວນວິຊາ — ນັບຈາກຕາຕະລາງວິຊາ");
        Bullet(body, "ຈຳນວນນັກຮຽນທີ່ມີວິຊາຕົກ — ໃນປີ + ພາກປະຈຸບັນ (ບໍ່ນັບຄຸນສົມບັດ ແລະ ການອອກແຮງງານ)");
        Bullet(body, "ກຣາຟແທ່ງສະແດງຈຳນວນນັກຮຽນແຍກຕາມຊັ້ນ (ມ.1 – ມ.4)");
        Bullet(body, "ປະກາດຂ່າວສານ 5 ລາຍການລ່າສຸດ");
        body.AppendChild(PageBreak());

        // ── 4. Students ───────────────────────────────────────
        H1(body, "4. ການຈັດການນັກຮຽນ");
        Body(body, "ໜ້າ \"ຂໍ້ມູນນັກຮຽນ\" ໃຊ້ສຳລັບເພີ່ມ, ແກ້ໄຂ ແລະ ລົບຂໍ້ມູນນັກຮຽນ. ມີຕົວກອງ ຊັ້ນ, ຫ້ອງ, ປີ, ສະຖານະ ແລະ ຊ່ອງຄົ້ນຫາ.");
        H2(body, "ເພີ່ມນັກຮຽນໃໝ່:");
        Step(body, "1.", "ກົດປຸ່ມ \"➕ ເພີ່ມນັກຮຽນ\" ດ້ານເທິງ.");
        Step(body, "2.", "ປ້ອນຂໍ້ມູນແບ່ງເປັນ 4 ແທັບ: ຂໍ້ມູນສ່ວນຕົວ, ທີ່ຢູ່, ຂໍ້ມູນພໍ່ແມ່, ຂໍ້ມູນການຮຽນ.");
        Step(body, "3.", "ຊ່ອງທີ່ມີ \"*\" ແມ່ນຕ້ອງປ້ອນ (ລະຫັດ, ຊື່, ນາມສະກຸນ, ຊັ້ນ, ປີການສຶກສາ).");
        Step(body, "4.", "ກົດ \"💾 ບັນທຶກ\" — ນັກຮຽນຈະປະກົດຢູ່ໃນຕາຕະລາງທັນທີ.");
        Tip(body, "ປີການສຶກສາ ໃນຟອມຈະຮັບເອົາລາຍຊື່ປີຈາກໜ້າ ‘ປີການສຶກສາ’ ໂດຍອັດຕະໂນມັດ.");

        H2(body, "ແກ້ໄຂນັກຮຽນ:");
        Step(body, "1.", "ເລືອກນັກຮຽນຈາກຕາຕະລາງ (ຫຼື double-click ແຖວນັ້ນ).");
        Step(body, "2.", "ກົດ \"✏️ ແກ້ໄຂ\" — ຟອມຈະປະກົດພ້ອມຂໍ້ມູນເດີມ.");
        Step(body, "3.", "ປ່ຽນແປງຕາມຕ້ອງການ ແລ້ວກົດ \"💾 ບັນທຶກ\".");

        H2(body, "ລົບນັກຮຽນ:");
        Step(body, "1.", "ເລືອກນັກຮຽນ → ກົດ \"🗑 ລົບ\".");
        Step(body, "2.", "ຢືນຢັນ. ການລົບຈະນຳເອົາການລົງທະບຽນ, ຄະແນນ ແລະ ປະຫວັດທັງໝົດອອກນຳ.");
        Tip(body, "ຄຳເຕືອນ: ການລົບແມ່ນກູ້ຄືນບໍ່ໄດ້. ຖ້າແມ່ນນັກຮຽນທີ່ຈົບການສຶກສາ ໃຫ້ປ່ຽນສະຖານະເປັນ ‘ຈົບ’ ແທນການລົບ.");
        body.AppendChild(PageBreak());

        // ── 5. Enrollment ─────────────────────────────────────
        H1(body, "5. ການລົງທະບຽນວິຊາ");
        Body(body, "ມີ 2 ໜ້າສຳລັບການລົງທະບຽນ:");
        Bullet(body, "ລົງທະບຽນ (ລາຍຄົນ) — ນັກຮຽນທີລະຄົນ");
        Bullet(body, "ລົງທະບຽນ (Batch) — ທັງຫ້ອງພ້ອມກັນ");

        H2(body, "ລົງທະບຽນ (ລາຍຄົນ):");
        Step(body, "1.", "ເລືອກ \"ປີ\", \"ຊັ້ນ\" ແລະ ນັກຮຽນ.");
        Step(body, "2.", "ກົດປຸ່ມ \"⚡ ລົງທະບຽນທຸກວິຊາ\".");
        Step(body, "3.", "ຢືນຢັນ — ລະບົບຈະລົງທະບຽນທຸກວິຊາທີ່ກຳນົດໄວ້ສຳລັບຊັ້ນນັ້ນ ທັງສອງພາກໂດຍອັດຕະໂນມັດ.");
        Step(body, "4.", "ປະກົດຂໍ້ຄວາມສະຫຼຸບ: ‘ບັນທຶກໃໝ່ N ວິຊາ, ມີຢູ່ແລ້ວ M ວິຊາ’.");
        Tip(body, "ກົດຊ້ຳໄດ້ — ການລົງທະບຽນວິຊາທີ່ມີຢູ່ແລ້ວຈະຖືກຂ້າມໂດຍອັດຕະໂນມັດ (ບໍ່ມີຂໍ້ມູນຊ້ຳ).");

        H2(body, "ລົງທະບຽນ (Batch):");
        Step(body, "1.", "ເລືອກ \"ປີ\", \"ຊັ້ນ\" ແລະ \"ຫ້ອງ\" (ຫຼື ‘ທັງໝົດ’).");
        Step(body, "2.", "ກວດເບິ່ງລາຍຊື່ນັກຮຽນທີ່ຈະຖືກລົງທະບຽນ.");
        Step(body, "3.", "ກົດ \"⚡ ລົງທະບຽນທຸກນັກຮຽນ\" → ຢືນຢັນ.");
        Step(body, "4.", "ລະບົບລົງທະບຽນຄົບທຸກນັກຮຽນ × ທຸກວິຊາ × ສອງພາກ ໃນຄຣັ້ງດຽວ.");

        H2(body, "ຖອນວິຊາ (ລາຍຄົນ):");
        Step(body, "1.", "ໃນໜ້າ ‘ລົງທະບຽນ (ລາຍຄົນ)’, ເລືອກວິຊາໃນຕາຕະລາງ.");
        Step(body, "2.", "ກົດ \"🗑 ຖອນວິຊາ\" → ຢືນຢັນ.");
        body.AppendChild(PageBreak());

        // ── 6. Monthly scores ─────────────────────────────────
        H1(body, "6. ການບັນທຶກຄະແນນປະຈຳເດືອນ");
        Body(body, "ໜ້າ ‘ບັນທຶກຄະແນນປະຈຳເດືອນ’ ໃຊ້ສຳລັບປ້ອນຄະແນນລາຍເດືອນຂອງແຕ່ລະວິຊາ. ສູດ:");
        Body(body, "    ກິດຈະກຳ (/3) + ວິໄນ (/2) + ການບ້ານ (/5) = ຄະແນນປະຈຳເດືອນ (/10)");
        H2(body, "ປະຕິທິນ:");
        Bullet(body, "ພາກ 1: ກ.ຍ. · ຕ.ລ. · ພ.ຍ. · ທ.ວ. (4 ເດືອນ) → ສອບປາຍພາກ ມ.ກ.");
        Bullet(body, "ພາກ 2: ກ.ພ. · ມີ.ນ. · ມ.ສ. · ພ.ພ. (4 ເດືອນ) → ສອບປາຍພາກ ມິ.ຖ.");
        H2(body, "ຂັ້ນຕອນ:");
        Step(body, "1.", "ເລືອກ \"ປະເພດ\" = \"ປະຈຳເດືອນ\".");
        Step(body, "2.", "ເລືອກ ຊັ້ນ, ຫ້ອງ, ສະຖານະ, ວິຊາ, ປີ ແລະ ເດືອນ.");
        Step(body, "3.", "ກົດ \"🔄 ໂຫຼດ\" — ຕາຕະລາງຂອງນັກຮຽນຈະປະກົດ.");
        Step(body, "4.", "ປ້ອນຄະແນນແຕ່ລະຊ່ອງ (ໃຊ້ dropdown ຈຳນວນເຕັມເທົ່ານັ້ນ).");
        Step(body, "5.", "ກົດ \"💾 ບັນທຶກ\" ຫຼື ໃຊ້ ໂຄ້ດ Ctrl+S.");
        Tip(body, "ສະເລ່ຍຂອງ 4 ເດືອນຈະຖືກຄຳນວນອັດຕະໂນມັດ ແລະ ສະແດງເປັນ ‘ສະເລ່ຍປະຈຳເດືອນ’ ໃນໜ້າ ‘ບັນທຶກຄະແນນປະຈຳພາກ’.");
        H2(body, "ໂໝດອ່ານຢ່າງດຽວ:");
        Body(body, "ຖ້າປ່ຽນ ‘ສະຖານະ’ ໄປ ‘ຈົບ’ ຫຼື ‘ທັງໝົດ’, ຕາຕະລາງຈະປ່ຽນເປັນອ່ານຢ່າງດຽວ — ສຳລັບດູປະຫວັດຂອງນັກຮຽນທີ່ຈົບແລ້ວ.");
        body.AppendChild(PageBreak());

        // ── 7. Semester scores ────────────────────────────────
        H1(body, "7. ການບັນທຶກຄະແນນປະຈຳພາກ");
        Body(body, "ໜ້າ ‘ບັນທຶກຄະແນນພາກຮຽນ’ ໃຊ້ສຳລັບປ້ອນຄະແນນສອບປາຍພາກ. ສູດສຳລັບຄະແນນລວມ:");
        Body(body, "    ຄະແນນລວມ = ສະເລ່ຍປະຈຳເດືອນ × 50% + ສອບເສງປາຍພາກ × 50%");
        H2(body, "ຂັ້ນຕອນ:");
        Step(body, "1.", "ເລືອກ ຊັ້ນ, ຫ້ອງ, ສະຖານະ, ນັກຮຽນ, ປີ ແລະ ພາກ.");
        Step(body, "2.", "ລາຍການວິຊາຂອງນັກຮຽນຈະປະກົດໃນຕາຕະລາງ ພ້ອມສະເລ່ຍປະຈຳເດືອນທີ່ໄດ້ມາໂດຍອັດຕະໂນມັດ.");
        Step(body, "3.", "ປ້ອນຄະແນນ \"ສອບເສງພາກຮຽນ (/10)\" — ໃຊ້ dropdown 0–10.");
        Step(body, "4.", "ຄະແນນລວມ ແລະ ລະດັບຈະຖືກຄຳນວນອັດຕະໂນມັດ.");
        Step(body, "5.", "ກົດ \"💾 ບັນທຶກ\" ຫຼື Ctrl+S.");
        H2(body, "ລະດັບ:");
        Bullet(body, "≥ 8.0 → ດີຫຼາຍ");
        Bullet(body, "6.0 – 7.99 → ດີ");
        Bullet(body, "5.0 – 5.99 → ຜ່ານ");
        Bullet(body, "< 5.0 → ບໍ່ຜ່ານ");
        H2(body, "ໝາຍເຫດ:");
        Body(body, "ວິຊາ ‘ຄຸນສົມບັດ’ ແລະ ‘ການອອກແຮງງານ’ ຈະບໍ່ປະກົດໃນໜ້ານີ້ ເພາະວ່າມັນບໍ່ມີສອບປາຍພາກ — ໃຫ້ປ້ອນຢູ່ໜ້າ ‘ບັນທຶກຄະແນນປະຈຳເດືອນ’ ໂດຍປ່ຽນ ‘ປະເພດ’ ເປັນ ‘ສະຫຼຸບພາກ 1/2/ປະຈຳປີ’ (ເບິ່ງບົດ 8).");
        body.AppendChild(PageBreak());

        // ── 8. CHA/LAB ────────────────────────────────────────
        H1(body, "8. ຄຸນສົມບັດ ແລະ ການອອກແຮງງານ");
        Body(body, "ຄຸນສົມບັດ (CHA1) ແລະ ການອອກແຮງງານ (LAB1) ເປັນວິຊາທີ່ຄະແນນຕ້ອງປ້ອນດ້ວຍຕົນເອງ — ບໍ່ມີການຄຳນວນອັດຕະໂນມັດ. ມີ 2 ປະເພດການປ້ອນ:");
        H2(body, "1) ປະຈຳເດືອນ:");
        Body(body, "ໃນໜ້າ ‘ບັນທຶກຄະແນນປະຈຳເດືອນ’ ປ້ອນຄະແນນແບບເດີມ (ກິດຈະກຳ + ວິໄນ + ການບ້ານ).");
        H2(body, "2) ສະຫຼຸບພາກ / ປະຈຳປີ:");
        Step(body, "1.", "ໃນໜ້າ ‘ບັນທຶກຄະແນນປະຈຳເດືອນ’, ປ່ຽນ ‘ປະເພດ’ ເປັນ ‘ສະຫຼຸບພາກຮຽນ 1’, ‘ສະຫຼຸບພາກຮຽນ 2’ ຫຼື ‘ສະຫຼຸບປະຈຳປີ’.");
        Step(body, "2.", "ວິຊາຈະຖືກກອງເຫຼືອພຽງ CHA1 ແລະ LAB1.");
        Step(body, "3.", "ປ້ອນຄະແນນ /10 ດ້ວຍ dropdown.");
        Step(body, "4.", "ກົດ \"💾 ບັນທຶກ\".");
        Tip(body, "ຄະແນນເຫຼົ່ານີ້ສະແດງໃນລາຍງານສະຫຼຸບພາກ ແລະ ປະຈຳປີໂດຍບໍ່ມີການສະເລ່ຍ — ປ້ອນແບບໃດປະກົດແບບນັ້ນ.");
        body.AppendChild(PageBreak());

        // ── 9. Reports ────────────────────────────────────────
        H1(body, "9. ການອອກລາຍງານ");
        Body(body, "ໜ້າ ‘ໃບຄະແນນ / ລາຍງານ’ ມີ 6 ປະເພດລາຍງານໃຫ້ເລືອກຈາກ dropdown ‘ປະເພດ’ ດ້ານເທິງ:");
        Bullet(body, "📅 ໃບຄະແນນປະຈຳເດືອນ (Excel) — ສຳລັບ 1 ຫ້ອງ × 1 ເດືອນ");
        Bullet(body, "📑 ໃບສັນຍາເຂົ້າຮຽນ (Word/PDF) — ສຳລັບນັກຮຽນທີລະຄົນ");
        Bullet(body, "📊 ສະຫຼຸບສະເລ່ຍພາກຮຽນ 1 (Excel) — ສຳລັບ 1 ຫ້ອງ");
        Bullet(body, "📊 ສະຫຼຸບສະເລ່ຍພາກຮຽນ 2 (Excel) — ສຳລັບ 1 ຫ້ອງ");
        Bullet(body, "📆 ໃບຄະແນນປະຈຳປີ (Excel) — ສະຫຼຸບສະເລ່ຍສອງພາກ");
        Bullet(body, "👤 ໃບຄະແນນບຸກຄົນ (Individual) — ປະຈຳເດືອນ / ສະຫຼຸບພາກ / ປະຈຳປີ ສຳລັບ 1 ນັກຮຽນ");
        H2(body, "ຂັ້ນຕອນທົ່ວໄປ:");
        Step(body, "1.", "ເລືອກ ‘ປະເພດ’ ລາຍງານ.");
        Step(body, "2.", "ເລືອກຕົວກອງທີ່ປະກົດ (ປີ, ພາກ, ຊັ້ນ, ຫ້ອງ, ນັກຮຽນ, ເດືອນ, ສະຖານະ — ຂຶ້ນກັບປະເພດລາຍງານ).");
        Step(body, "3.", "ກວດເບິ່ງ preview ໃນແທັບ ‘📊 ຕາຕະລາງ’.");
        Step(body, "4.", "ກົດ \"📋 ສ້າງ Excel\" ຫຼື \"📄 ສ້າງ PDF\" ເພື່ອບັນທຶກໄຟລ໌.");
        H2(body, "ຄຸນລັກສະນະພິເສດ:");
        Bullet(body, "ການແປງ PDF ໃຊ້ Microsoft Excel / Word ທີ່ຕິດຕັ້ງໃນເຄື່ອງ — PDF ຈະມີຮູບແບບຄືກັນກັບ Excel/Word ທຸກປະການ.");
        Bullet(body, "ສຳລັບລາຍງານ Profile ແລະ Enrollment Agreement, ສາມາດເລືອກບັນທຶກເປັນ .docx ຫຼື .pdf ໃນ Save dialog.");
        Bullet(body, "ນັກຮຽນທີ່ຈົບແລ້ວສາມາດດຶງ ‘ໃບຄະແນນບຸກຄົນ’ ໂດຍປ່ຽນ ‘ສະຖານະ’ ເປັນ ‘ຈົບ’ ໃນຕົວກອງ.");
        Tip(body, "ຖ້າຕົວອັກສອນລາວສະແດງເປັນກ່ອງສີ່ຫຼ່ຽມໃນ PDF, ໃຫ້ຕິດຕັ້ງຟອນ ‘Phetsarath OT’ ໃນເຄື່ອງຄອມພິວເຕີ.");
        body.AppendChild(PageBreak());

        // ── 10. Promotion ─────────────────────────────────────
        H1(body, "10. ການຂຶ້ນຊັ້ນ ແລະ ການຈົບການສຶກສາ");
        Body(body, "ໜ້າ ‘ຂຶ້ນຊັ້ນ / ຈົບ’ ໃຊ້ສຳລັບການເຄື່ອນຍ້າຍນັກຮຽນໄປຍັງ ປີ-ຊັ້ນ ໃໝ່ໃນຄຣັ້ງດຽວ.");
        H2(body, "3 ການກະທຳ:");
        Bullet(body, "ຂຶ້ນຊັ້ນ — ມ.1 → ມ.2, ມ.2 → ມ.3, ມ.3 → ມ.4, ມ.4 → ‘ຈົບ’");
        Bullet(body, "ຄືນຊັ້ນ (ຮຽນຊໍ້າ) — ຍັງຢູ່ຊັ້ນເດີມໃນປີໃໝ່");
        Bullet(body, "ຈົບການສຶກສາ — ປ່ຽນສະຖານະເປັນ ‘ຈົບ’ ໂດຍກົງ");
        H2(body, "ຂັ້ນຕອນ:");
        Step(body, "1.", "ເລືອກ ‘ປີ’ ປະຈຸບັນ ແລະ ‘ຊັ້ນ’ ທີ່ຕ້ອງການເຄື່ອນຍ້າຍ.");
        Step(body, "2.", "ປ້ອນ ‘ປີໃໝ່’ (ຮູບແບບ YYYY-YYYY ເຊັ່ນ 2026-2027).");
        Step(body, "3.", "ເລືອກ ‘ການກະທຳ’.");
        Step(body, "4.", "ໝາຍຖືກນັກຮຽນທີ່ຕ້ອງການ — ຫຼື ‘ເລືອກທັງໝົດ’.");
        Step(body, "5.", "ກົດ \"⚡ ດຳເນີນການ\" → ຢືນຢັນ.");
        Tip(body, "ຂໍ້ມູນເດີມຍັງຄົງຢູ່ — ການຂຶ້ນຊັ້ນແມ່ນການປ່ຽນ GradeLevel/AcademicYear ຂອງນັກຮຽນ ບໍ່ແມ່ນລົບ. ປະຫວັດການຂຶ້ນຊັ້ນເກັບໃນຕາຕະລາງ ‘ປະຫວັດການຂຶ້ນຊັ້ນ’.");
        body.AppendChild(PageBreak());

        // ── 11. Academic Year ─────────────────────────────────
        H1(body, "11. ການຈັດການປີການສຶກສາ");
        Body(body, "ໜ້າ ‘ປີການສຶກສາ’ (ສຳລັບ admin) ໃຊ້ສຳລັບເພີ່ມ, ປ່ຽນ ແລະ ລົບປີການສຶກສາ.");
        H2(body, "ການກະທຳຫຼັກ:");
        Bullet(body, "➕ ເພີ່ມປີໃໝ່ — ປ້ອນຮູບແບບ YYYY-YYYY");
        Bullet(body, "⚡ ຂຶ້ນປີໃໝ່ອັດຕະໂນມັດ — ປະຈຸບັນ+1 ປີ, ກັບໄປພາກ 1");
        Bullet(body, "✅ ຕັ້ງເປັນປະຈຸບັນ — ປ່ຽນປີຫຼັກ");
        Bullet(body, "🗑 ລຶບປີ (ບໍ່ມີຂໍ້ມູນ) — ປອດໄພ; ປະຕິເສດຖ້າຍັງມີຂໍ້ມູນ");
        Bullet(body, "🔥 ບັງຄັບລຶບ + ຂໍ້ມູນທັງໝົດ — ລົບປີພ້ອມຂໍ້ມູນທັງໝົດ (ກູ້ຄືນບໍ່ໄດ້; ຕ້ອງພິມຊື່ປີຢືນຢັນ)");
        Tip(body, "ບໍ່ສາມາດລົບປີທີ່ເປັນ ‘ປະຈຸບັນ’ ໄດ້ — ໃຫ້ປ່ຽນປະຈຸບັນຫາປີອື່ນກ່ອນ.");
        body.AppendChild(PageBreak());

        // ── 12. Subjects + Users ──────────────────────────────
        H1(body, "12. ການຈັດການວິຊາ ແລະ ຜູ້ໃຊ້ລະບົບ");
        H2(body, "ໜ້າ ‘ວິຊາ’:");
        Body(body, "ສະແດງລາຍຊື່ວິຊາທີ່ໃຊ້ໃນລະບົບ (14 ວິຊາມາດຕະຖານ MoES). ສາມາດເພີ່ມ, ແກ້ໄຂ ຫຼື ລົບໄດ້. ໂດຍປົກກະຕິບໍ່ຄວນແກ້ໄຂລາຍຊື່ນີ້ — ມັນເປັນພື້ນຖານຂອງລະບົບການຄິດໄລ່.");
        H2(body, "ໜ້າ ‘ຈັດການຜູ້ໃຊ້’ (admin ເທົ່ານັ້ນ):");
        Body(body, "ສຳລັບເພີ່ມ, ແກ້ໄຂ ແລະ ລົບບັນຊີຜູ້ໃຊ້ລະບົບ. ມີ 2 ບົດບາດ:");
        Bullet(body, "admin — ໃຊ້ໄດ້ທຸກໜ້າ");
        Bullet(body, "teacher — ບໍ່ມີສິດເຂົ້າເຖິງ ‘ຂຶ້ນຊັ້ນ’, ‘ຈັດການຜູ້ໃຊ້’, ‘ປີການສຶກສາ’, ‘ຕັ້ງຄ່າລະບົບ’");
        H2(body, "ການປ່ຽນລະຫັດຜ່ານ:");
        Body(body, "Admin ສາມາດແກ້ໄຂລະຫັດຜ່ານຂອງຜູ້ໃຊ້ໃດກໍ່ໄດ້ ໃນໜ້າ ‘ຈັດການຜູ້ໃຊ້’ — ດັບເບີນຄຣິກຜູ້ໃຊ້ ແລ້ວປ້ອນລະຫັດໃໝ່.");
        body.AppendChild(PageBreak());

        // ── 13. Settings + Backup ─────────────────────────────
        H1(body, "13. ການຕັ້ງຄ່າລະບົບ / Backup ແລະ Restore");
        Body(body, "ໜ້າ ‘ຕັ້ງຄ່າລະບົບ’ ມີ 4 ແທັບ:");
        H2(body, "ແທັບ ‘ທົ່ວໄປ’:");
        Body(body, "ປ່ຽນຊື່ໂຮງຮຽນ. ປີການສຶກສາ ແລະ ພາກຮຽນ ຈັດການແຍກໃນໜ້າ ‘ປີການສຶກສາ’.");
        H2(body, "ແທັບ ‘ນຳເຂົ້າ CSV’:");
        Step(body, "1.", "ກົດ \"📄 Template\" ເພື່ອດາວໂຫຼດແບບຟອມ CSV (28 ຄໍລຳ).");
        Step(body, "2.", "ກອກຂໍ້ມູນນັກຮຽນໃນ Excel ບັນທຶກເປັນ CSV (UTF-8).");
        Step(body, "3.", "ກົດ \"📂 ເລືອກໄຟລ໌ ແລະ ນຳເຂົ້າ\" → ເລືອກໄຟລ໌ CSV.");
        Step(body, "4.", "ສະຫຼຸບຈະປະກົດ: ‘ນຳເຂົ້າ N, ຂ້າມ M (ລະຫັດຊ້ຳ), ບໍ່ຄົບ K (ຂາດຂໍ້ມູນຈຳເປັນ)’.");
        H2(body, "ແທັບ ‘Backup’:");
        Body(body, "Backup: ສຳເນົາໄຟລ໌ sis_lao.db ໄປໄວ້ຍັງທີ່ປອດໄພ (USB / cloud / network).");
        Body(body, "Restore: ນຳໄຟລ໌ backup ກັບມາ — ຈະທົດແທນຂໍ້ມູນທັງໝົດທີ່ມີຢູ່!");
        Tip(body, "ຄຳແນະນຳ: Backup ປະຈຳເດືອນຢ່າງໜ້ອຍເດືອນລະ 1 ຄຣັ້ງ ແລະ ກ່ອນເຮັດການກະທຳໃຫຍ່ໆ (ຂຶ້ນຊັ້ນ, ບັງຄັບລຶບປີ, ນຳເຂົ້າ CSV).");
        H2(body, "ແທັບ ‘Log’:");
        Body(body, "ສະແດງປະຫວັດການກະທຳຂອງລະບົບ (login, ບັນທຶກຄະແນນ, ສ້າງລາຍງານ, backup, restore...). ສາມາດລົບ log ທີ່ເກົ່າກວ່າ 30 ວັນໄດ້.");
        body.AppendChild(PageBreak());

        // ── 14. Troubleshooting ──────────────────────────────
        H1(body, "14. ການແກ້ໄຂບັນຫາທີ່ພົບເລື້ອຍ");
        H2(body, "ບັນຊີ admin ບໍ່ສາມາດເຂົ້າສູ່ລະບົບໄດ້:");
        Body(body, "ກວດສອບການພິມຕົວອັກສອນ. Username ແລະ Password ກ່ຽວຂ້ອງກັບຕົວເລັກ-ໃຫຍ່. ຖ້າລືມລະຫັດຜ່ານ admin ຕ້ອງ Restore ຈາກ backup ທີ່ມີລະຫັດເກົ່າ.");
        H2(body, "ບໍ່ສາມາດສ້າງ PDF ໄດ້ (Excel COM error):");
        Body(body, "ການແປງ PDF ຕ້ອງມີ Microsoft Excel ຫຼື Word ຕິດຕັ້ງໃນເຄື່ອງ. ຖ້າບໍ່ມີ, ໃຫ້ບັນທຶກເປັນ Excel/Word ກ່ອນ ແລ້ວໄປແປງ PDF ໃນ Office ໂດຍກົງ.");
        H2(body, "ຕົວອັກສອນລາວເປັນກ່ອງສີ່ຫຼ່ຽມໃນລາຍງານ:");
        Body(body, "ເຄື່ອງບໍ່ມີຟອນ Phetsarath OT ຕິດຕັ້ງ. ດາວໂຫຼດແລະຕິດຕັ້ງ ‘Install for all users’. ລະບົບຈະນຳໃຊ້ໂດຍອັດຕະໂນມັດໃນຄຣັ້ງຕໍ່ໄປ.");
        H2(body, "ບໍ່ສາມາດແກ້ໄຂຄະແນນຂອງນັກຮຽນຈົບແລ້ວ:");
        Body(body, "ໂດຍການອອກແບບ — ນັກຮຽນທີ່ ‘ຈົບ’ ມີຂໍ້ມູນອ່ານຢ່າງດຽວ. ຖ້າຕ້ອງການແກ້ໄຂ, ໃຫ້ປ່ຽນສະຖານະຊົ່ວຄາວເປັນ ‘ກຳລັງຮຽນ’ ໃນໜ້າ ‘ຂໍ້ມູນນັກຮຽນ’ ກ່ອນ.");
        H2(body, "ຂໍ້ມູນເສຍ ຫຼື ໄຟລ໌ sis_lao.db ເສຍ:");
        Step(body, "1.", "ປິດໂປຣແກຣມ.");
        Step(body, "2.", "ນຳໄຟລ໌ backup .db ມາ.");
        Step(body, "3.", "ເປີດໂປຣແກຣມ → Settings → Backup → \"🔄 Restore\" → ເລືອກໄຟລ໌ backup.");
        Step(body, "4.", "ໂປຣແກຣມຈະປິດແລະເປີດໃໝ່ໂດຍອັດຕະໂນມັດ.");

        // ── Back cover ────────────────────────────────────────
        body.AppendChild(PageBreak());
        for (int i = 0; i < 8; i++) body.AppendChild(BlankLine());
        body.AppendChild(P("— ສິ້ນສຸດຄູ່ມືການໃຊ້ງານ —", size: 28, bold: true, center: true));
        body.AppendChild(BlankLine());
        body.AppendChild(P("StudentSIS Lao  ·  ໂຮງຮຽນ ມສ ບຶກທົ່ງ",
            size: 24, center: true, color: "4B5563"));

        // Set page size + margins on the section.
        body.AppendChild(new SectionProperties(
            new PageSize { Width = 12240U, Height = 15840U },           // A4 (twips)
            new PageMargin {
                Top = 1134, Right = 1134, Bottom = 1134, Left = 1134,
                Header = 720, Footer = 720, Gutter = 0
            }
        ));

        mainPart.Document.Save();
    }

    // ─────────────────────────────────────────────────────────────
    //  OpenXml authoring helpers
    // ─────────────────────────────────────────────────────────────
    private static Paragraph P(string text, int size = 22, bool bold = false,
        bool italic = false, bool center = false, string? color = null)
    {
        var runProps = new RunProperties();
        // Default Lao font + ASCII/EastAsian/ComplexScript fallback so Word picks
        // the right typeface for the right script range automatically.
        runProps.Append(new RunFonts {
            Ascii = LaoFontName, HighAnsi = LaoFontName,
            ComplexScript = LaoFontName, EastAsia = LaoFontName
        });
        runProps.Append(new FontSize { Val = size.ToString() });
        runProps.Append(new FontSizeComplexScript { Val = size.ToString() });
        if (bold)
        {
            runProps.Append(new Bold());
            runProps.Append(new BoldComplexScript());
        }
        if (italic)
        {
            runProps.Append(new Italic());
            runProps.Append(new ItalicComplexScript());
        }
        if (color != null)
            runProps.Append(new Color { Val = color });

        var paraProps = new ParagraphProperties();
        if (center) paraProps.Append(new Justification { Val = JustificationValues.Center });
        paraProps.Append(new SpacingBetweenLines { After = "120", Line = "300", LineRule = LineSpacingRuleValues.Auto });

        var run = new Run(runProps, new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return new Paragraph(paraProps, run);
    }

    private static Paragraph BlankLine() => P("", size: 22);

    private static Paragraph PageBreak()
    {
        var run = new Run(new Break { Type = BreakValues.Page });
        return new Paragraph(run);
    }

    private static void H1(Body body, string text)
        => body.AppendChild(P(text, size: 36, bold: true, color: "1B4F8A"));

    private static void H2(Body body, string text)
        => body.AppendChild(P(text, size: 26, bold: true, color: "1E293B"));

    private static void Body(Body body, string text)
        => body.AppendChild(P(text, size: 22));

    private static void Bullet(Body body, string text)
        => body.AppendChild(P("   •  " + text, size: 22));

    private static void Step(Body body, string num, string text)
        => body.AppendChild(P($"   {num}  {text}", size: 22));

    private static void Tip(Body body, string text)
        => body.AppendChild(P("💡  " + text, size: 20, italic: true, color: "6B7280"));

    // ─────────────────────────────────────────────────────────────
    //  .docx → .pdf converters (Word COM and LibreOffice)
    // ─────────────────────────────────────────────────────────────
    private static bool TryConvertViaWord(string docx, string pdf, out string error)
    {
        error = "";
        try
        {
            Type? wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null) { error = "Word.Application ProgID not found"; return false; }
            dynamic? word = Activator.CreateInstance(wordType);
            if (word == null) { error = "could not create Word.Application instance"; return false; }
            try
            {
                word.Visible = false;
                word.DisplayAlerts = 0;
                dynamic docs = word.Documents;
                dynamic d = docs.Open(Path.GetFullPath(docx),
                    /*ConfirmConversions*/false, /*ReadOnly*/true,
                    /*AddToRecentFiles*/false);
                // 17 = wdFormatPDF
                d.SaveAs2(Path.GetFullPath(pdf), 17);
                d.Close(false);
                return true;
            }
            finally { try { word.Quit(); } catch { } }
        }
        catch (Exception ex) { error = ex.Message; return false; }
    }

    private static bool TryConvertViaLibreOffice(string docx, string pdf, out string error)
    {
        error = "";
        string[] candidates = {
            @"C:\Program Files\LibreOffice\program\soffice.exe",
            @"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        };
        string? soffice = null;
        foreach (var c in candidates) if (File.Exists(c)) { soffice = c; break; }
        if (soffice == null) { error = "soffice.exe not found"; return false; }
        try
        {
            string outDir = Path.GetDirectoryName(Path.GetFullPath(pdf))!;
            var psi = new ProcessStartInfo(soffice,
                $"--headless --convert-to pdf --outdir \"{outDir}\" \"{docx}\"")
            { CreateNoWindow = true, UseShellExecute = false, RedirectStandardOutput = true };
            using var p = Process.Start(psi);
            if (p == null) { error = "Process.Start returned null"; return false; }
            p.WaitForExit(60_000);
            // LibreOffice writes <basename>.pdf in outdir.
            string libOut = Path.Combine(outDir, Path.GetFileNameWithoutExtension(docx) + ".pdf");
            if (!File.Exists(libOut)) { error = "LibreOffice produced no PDF"; return false; }
            if (!string.Equals(libOut, Path.GetFullPath(pdf), StringComparison.OrdinalIgnoreCase))
            {
                File.Copy(libOut, pdf, true);
                try { File.Delete(libOut); } catch { }
            }
            return true;
        }
        catch (Exception ex) { error = ex.Message; return false; }
    }
}
