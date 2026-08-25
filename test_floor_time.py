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

    def test_monthly_roster_image_grows_with_agent_count(self):
        small_agents = self.make_agents(3)
        large_agents = self.make_agents(40)
        small = app.render_monthly_agent_roster_image(
            small_agents, 2026, 8, ["Nama", "Unit", "Total Kehadiran"],
            "Default (sesuai file)", "HUB CIBUBUR", "Brighton",
        )
        large = app.render_monthly_agent_roster_image(
            large_agents, 2026, 8, ["Nama", "Unit", "Total Kehadiran"],
            "Default (sesuai file)", "HUB CIBUBUR", "Brighton",
        )
        self.assertGreater(large.height, small.height)
        self.assertEqual(small.width, large.width)


if __name__ == '__main__':
    unittest.main()
