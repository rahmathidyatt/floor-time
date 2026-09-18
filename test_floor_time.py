import sys
import types
import unittest
from datetime import date

try:
    import streamlit  # noqa: F401
except ImportError:
    sys.modules['streamlit'] = types.ModuleType('streamlit')

import app
import pandas as pd


class FloorTimeTests(unittest.TestCase):
    def make_agents(self, count):
        return [
            app.AgentRecord(
                agent_id=f"a{i}",
                name=f"AGENT {i}",
                code=f"C{i:03d}",
                business_unit="Paradise" if i % 2 else "Cibubur",
                previous_month_attendance=10 + i,
            )
            for i in range(count)
        ]

    def test_cross_month_week(self):
        dates = app.get_month_work_dates(2026, 9)
        self.assertEqual(dates[0], date(2026, 8, 31))
        self.assertEqual(dates[-1], date(2026, 10, 3))
        weeks = app.group_dates_by_calendar_week(dates)
        self.assertEqual(weeks[1][0], date(2026, 8, 31))
        self.assertEqual(weeks[1][-1], date(2026, 9, 5))

    def test_previous_month_cross_year(self):
        self.assertEqual(app.previous_month(2027, 1), (2026, 12))

    def test_excel_style_columns_are_normalized(self):
        df = pd.DataFrame([
            {
                "No": 1,
                "Nama": "TONI (ACYS)",
                "Jabatan": "Business Manager",
                "Office": "Brighton Priority Cibubur, Bogor",
                "Total": 13,
            }
        ])
        agents, warnings = app.dataframe_to_agents(df)
        self.assertEqual(warnings, [])
        self.assertEqual(len(agents), 1)
        self.assertEqual(agents[0].name, "TONI")
        self.assertEqual(agents[0].code, "ACYS")
        self.assertEqual(agents[0].job_title, "Business Manager")
        self.assertEqual(agents[0].office, "Brighton Priority Cibubur, Bogor")
        self.assertEqual(agents[0].previous_month_attendance, 13)

    def test_attendance_filter_and_manual_exclusion(self):
        agents = self.make_agents(6)
        # Attendance make_agents = 10..15. Threshold 13 keeps a3, a4, a5.
        filtered = app.filter_agents_for_schedule(agents, minimum_attendance=13)
        self.assertEqual([a.agent_id for a in filtered], ["a3", "a4", "a5"])

        filtered = app.filter_agents_for_schedule(
            agents,
            minimum_attendance=13,
            excluded_agent_ids={"a4"},
        )
        self.assertEqual([a.agent_id for a in filtered], ["a3", "a5"])

    def test_holiday_has_no_assignments_and_no_same_date_duplicate(self):
        dates = app.get_month_work_dates(2026, 9)
        weeks = app.group_dates_by_calendar_week(dates)
        holiday_date = date(2026, 9, 1)
        holidays = {holiday_date: app.HolidayInfo(holiday_date, "Maulid Nabi Muhammad SAW")}
        slots = app.build_slots_by_week(weeks, holidays, 2, 2)
        agents = self.make_agents(12)
        df, schedule, _, _ = app.generate_schedule(agents, slots, [], 1, True, "test-seed")
        self.assertFalse((df["Tanggal"] == holiday_date.isoformat()).any())
        duplicates = df.groupby(["Agent ID", "Tanggal"]).size()
        self.assertTrue((duplicates <= 1).all())
        ok, messages = app.validate_schedule(df, agents, weeks, holidays, 1)
        self.assertTrue(ok, messages)

    def test_small_agent_pool_can_repeat_on_different_days(self):
        dates = app.get_month_work_dates(2026, 9)
        weeks = app.group_dates_by_calendar_week(dates)
        slots = app.build_slots_by_week(weeks, {}, 2, 2)
        agents = self.make_agents(4)
        df, _, _, _ = app.generate_schedule(agents, slots, [], 3, False, "repeat-seed")
        first_week = df[df["Minggu Internal"] == 1]
        counts = first_week.groupby("Agent ID").size()
        self.assertTrue((counts <= 3).all())
        self.assertTrue((counts > 1).any())
        same_date = first_week.groupby(["Agent ID", "Tanggal"]).size()
        self.assertTrue((same_date <= 1).all())

    def test_insufficient_capacity_leaves_slots_unfilled_without_duplicate(self):
        dates = app.get_month_work_dates(2026, 9)
        weeks = app.group_dates_by_calendar_week(dates)
        slots = app.build_slots_by_week(weeks, {}, 4, 3)
        agents = self.make_agents(2)
        df, _, _, warnings = app.generate_schedule(agents, slots, [], 1, False, "small")
        self.assertTrue(any("belum terisi" in w for w in warnings))
        duplicate = df.groupby(["Agent ID", "Tanggal"]).size()
        self.assertTrue((duplicate <= 1).all())

    def test_export_font_is_scalable(self):
        from PIL import Image, ImageDraw, ImageFont

        font_small = app.find_inter_font(16, bold=False)
        font_large = app.find_inter_font(64, bold=True)
        canvas = Image.new("RGB", (500, 200), "white")
        draw = ImageDraw.Draw(canvas)
        small_w, small_h = app.text_size(draw, "Brighton", font_small)
        large_w, large_h = app.text_size(draw, "Brighton", font_large)
        self.assertGreater(large_w, small_w * 2)
        self.assertGreater(large_h, small_h * 2)
        self.assertIsInstance(font_large, (ImageFont.FreeTypeFont, ImageFont.ImageFont))

    def test_dynamic_image_grows_when_agent_list_is_large(self):
        week_dates = [date(2026, 8, 31), date(2026, 9, 1), date(2026, 9, 2), date(2026, 9, 3), date(2026, 9, 4), date(2026, 9, 5)]
        weeks = {1: week_dates}
        slots = app.build_slots_by_week(weeks, {}, 1, 1)[1]
        small_schedule = {s.key: [app.AssignmentEntry("a", "AGENT A (AAAA)")] for s in slots}
        large_schedule = dict(small_schedule)
        first_slot = slots[0]
        large_schedule[first_slot.key] = [app.AssignmentEntry(f"a{i}", f"AGENT PANJANG NOMOR {i} (C{i:02d})") for i in range(30)]
        small = app.render_week_image(week_dates, slots, small_schedule, {}, "HUB CIBUBUR", "Brighton", ["Catatan"], "Portrait")
        large = app.render_week_image(week_dates, slots, large_schedule, {}, "HUB CIBUBUR", "Brighton", ["Catatan"], "Portrait")
        self.assertGreater(large.height, small.height)

    def test_week_poster_notes_height_follows_content(self):
        week_dates = [date(2026, 9, 7), date(2026, 9, 8)]
        weeks = {1: week_dates}
        slots = app.build_slots_by_week(weeks, {}, 1, 1)[1]
        schedule = {s.key: [app.AssignmentEntry("a", "AGENT A (AAAA)")] for s in slots}
        short = app.render_week_image(
            week_dates, slots, schedule, {}, "HUB CIBUBUR", "Brighton",
            ["Catatan pertama."], "Portrait",
        )
        long = app.render_week_image(
            week_dates, slots, schedule, {}, "HUB CIBUBUR", "Brighton",
            ["Catatan pertama.", "Catatan kedua.", "Catatan ketiga.", "Catatan keempat."], "Portrait",
        )
        self.assertGreater(long.height, short.height)

    def test_week_poster_has_no_corner_or_bottom_bar(self):
        week_dates = [date(2026, 9, 7)]
        weeks = {1: week_dates}
        slots = app.build_slots_by_week(weeks, {}, 1, 1)[1]
        schedule = {s.key: [app.AssignmentEntry("a", "AGENT A (AAAA)")] for s in slots}
        image = app.render_week_image(
            week_dates, slots, schedule, {}, "HUB CIBUBUR", "Brighton", ["Catatan"], "Portrait",
        )
        self.assertEqual(image.getpixel((0, 0)), (255, 255, 255))
        self.assertEqual(image.getpixel((0, image.height - 1)), (255, 255, 255))

    def test_unit_is_second_word_from_office(self):
        agent = app.AgentRecord(
            agent_id="u1",
            name="TONI",
            code="ACYS",
            business_unit="Brighton Priority Cibubur, Bogor",
            previous_month_attendance=13,
            job_title="Business Manager",
        )
        self.assertEqual(agent.unit, "Priority")

    def test_roster_attendance_sort_modes(self):
        agents = [
            app.AgentRecord("a1", "A", "A1", "Brighton Priority Cibubur, Bogor", 9),
            app.AgentRecord("a2", "B", "B1", "Brighton Warrior Cibubur, Bogor", 19),
            app.AgentRecord("a3", "C", "C1", "Brighton Champion Cibubur, Bogor", 5),
        ]
        self.assertEqual(
            [a.previous_month_attendance for a in app.sort_agents_for_roster(agents, "Default (sesuai file)")],
            [9, 19, 5],
        )
        self.assertEqual(
            [a.previous_month_attendance for a in app.sort_agents_for_roster(agents, "Terbanyak ke terendah")],
            [19, 9, 5],
        )
        self.assertEqual(
            [a.previous_month_attendance for a in app.sort_agents_for_roster(agents, "Terendah ke terbanyak")],
            [5, 9, 19],
        )

    def test_monthly_roster_is_paginated_for_50_agents(self):
        pages = app.render_monthly_agent_roster_images(
            self.make_agents(50), 2026, 8, ["Nama", "Unit", "Total Kehadiran"],
            "Default (sesuai file)", "HUB CIBUBUR", "Brighton",
        )
        self.assertEqual(len(pages), 3)
        self.assertTrue(all(page.width == 1080 for page in pages))
        self.assertGreater(pages[0].height, pages[-1].height)

    def test_monthly_roster_keeps_one_page_for_small_list(self):
        pages = app.render_monthly_agent_roster_images(
            self.make_agents(3), 2026, 8, ["Nama", "Unit", "Total Kehadiran"],
            "Default (sesuai file)", "HUB CIBUBUR", "Brighton",
        )
        self.assertEqual(len(pages), 1)
        self.assertEqual(pages[0].width, 1080)
        self.assertEqual(pages[0].getpixel((0, pages[0].height - 1)), (255, 255, 255))

    def test_notes_match_user_input_even_on_saturday(self):
        supplied = ["Catatan operasional", "Penanganan klien sesuai absensi."]
        self.assertEqual(app.poster_notes(supplied, True), supplied)
        self.assertEqual(app.poster_notes([], True), [])
        self.assertEqual(app.poster_notes(["  "], False), [])

    def test_dense_shift_does_not_reflow_into_narrow_columns(self):
        from PIL import Image, ImageDraw
        draw = ImageDraw.Draw(Image.new("RGB", (1080, 200)))
        font = app.find_inter_font(app.WEEK_AGENT_FONT_SIZE, bold=True)
        entries = [app.AssignmentEntry(str(i), "AGENT PANJANG (ABCD)") for i in range(30)]
        heights = [app.agent_entries_height(draw, entries[:n], 300, font, 1) for n in (4, 8, 9, 20, 30)]
        self.assertEqual(heights, sorted(set(heights)))

    def test_every_agent_gets_exact_weekly_target(self):
        for count in (4, 24, 49, 60):
            for target in (1, 2, 3):
                with self.subTest(count=count, target=target):
                    agents = self.make_agents(count)
                    weeks = app.group_dates_by_calendar_week(app.get_month_work_dates(2026, 9))
                    slots = app.build_slots_by_week(weeks, {}, 4, 3)
                    df, schedule, capacities, _ = app.generate_schedule(agents, slots, [], target, True, "exact-target")
                    counts = df.groupby(["Minggu Internal", "Agent ID"]).size()
                    self.assertEqual(len(counts), len(weeks) * count)
                    self.assertTrue((counts == target).all())
                    self.assertTrue((df.groupby(["Agent ID", "Tanggal"]).size() == 1).all())
                    for week, week_slots in schedule.items():
                        for key, entries in week_slots.items():
                            self.assertLessEqual(len(entries), capacities[week][key])

    def test_target_limited_by_active_days(self):
        days = [date(2026, 9, 7), date(2026, 9, 8)]
        slots = app.build_slots_by_week({1: days}, {}, 1, 1)
        df, _, _, warnings = app.generate_schedule(self.make_agents(49), slots, [], 3, True, "short-week")
        self.assertTrue((df.groupby("Agent ID").size() == 2).all())
        self.assertTrue(any("hanya 2 hari aktif" in w for w in warnings))

    def test_urgent_counts_towards_target(self):
        days = app.get_month_work_dates(2026, 9)
        weeks = app.group_dates_by_calendar_week(days)
        slots = app.build_slots_by_week(weeks, {}, 1, 1)
        requests = [app.UrgentRequest("a0", weeks[1][0], 0)]
        df, _, _, _ = app.generate_schedule(self.make_agents(49), slots, requests, 3, True, "urgent")
        self.assertTrue((df.groupby(["Minggu Internal", "Agent ID"]).size() == 3).all())
        forced = df[(df["Agent ID"] == "a0") & (df["Tanggal"] == weeks[1][0].isoformat())]
        self.assertEqual(len(forced), 1)
        self.assertEqual(forced.iloc[0]["Urgent"], "Ya")

    def test_bundled_fonts_work_without_system_fonts(self):
        import os
        from unittest.mock import patch
        from PIL import ImageFont
        original = ImageFont.truetype
        def bundled_only(path, *args, **kwargs):
            if "/assets/fonts/" not in str(path).replace("\\", "/"):
                raise OSError("System fonts unavailable")
            return original(path, *args, **kwargs)
        app._FONT_CACHE.clear()
        app._FONT_SOURCE_CACHE.clear()
        previous = os.getcwd()
        try:
            os.chdir(__import__("tempfile").gettempdir())
            with patch.object(app.ImageFont, "truetype", side_effect=bundled_only):
                regular = app.find_inter_font(24, False)
                bold = app.find_inter_font(24, True)
                self.assertEqual(regular.getname()[1], "Book")
                self.assertEqual(bold.getname()[1], "Bold")
                self.assertNotEqual(bytes(regular.getmask("AGUNG (WRWZ)")), bytes(bold.getmask("AGUNG (WRWZ)")))
        finally:
            os.chdir(previous)

    def test_missing_font_has_actionable_error(self):
        from unittest.mock import patch
        app._FONT_CACHE.clear()
        with patch.object(app.ImageFont, "truetype", side_effect=OSError("missing")):
            with self.assertRaisesRegex(RuntimeError, "assets/fonts"):
                app.find_inter_font(24, True)

    def test_long_names_stay_on_one_line_at_full_font_size(self):
        from PIL import Image, ImageDraw
        days = [date(2026, 9, 5)]
        slots = app.build_slots_by_week({1: days}, {}, 4, 4)[1]
        for orientation in ("Portrait", "Landscape"):
            for name in ("YANTI ANGELA (NNZA)", "NAMA AGEN YANG SANGAT PANJANG SEKALI (ABCD)"):
                schedule = {slot.key: [app.AssignmentEntry("a", name)] for slot in slots}
                m = app.calculate_week_image_metrics(days, slots, schedule, {}, [], orientation)
                scale = m["scale"]
                text_width = (m["width"] - 2*m["margin"] - m["day_col_w"])//2 - 78*scale
                font = app.find_inter_font(app.WEEK_AGENT_FONT_SIZE*scale, True)
                draw = ImageDraw.Draw(Image.new("RGB", (1,1)))
                self.assertEqual(app.wrap_agent_name(draw, name, font, text_width), [name])
                self.assertEqual(font.size, 24*scale)


if __name__ == '__main__':
    unittest.main()
