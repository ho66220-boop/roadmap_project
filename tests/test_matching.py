# -*- coding: utf-8 -*-
import unittest
from pathlib import Path

from roadmap_engine.recommender import DataStore, RecommendationEngine, Offering, UniTrackRow, norm


def make_engine(depts=(), master=(), schools=()):
    engine = RecommendationEngine.__new__(RecommendationEngine)
    engine.store = DataStore(
        dept_recs={dept: {} for dept in depts},
        subject_master={norm(name): name for name in master},
        schools={school: [] for school in schools},
    )
    return engine


class MajorMatchingTest(unittest.TestCase):
    def test_pick_major_prefers_exact_before_abbreviation_replacement(self):
        engine = make_engine(["생명과학", "생명공학", "컴퓨터교육", "컴퓨터공학"])
        self.assertEqual(engine._pick_major("생명과학"), "생명과학")
        self.assertEqual(engine._pick_major("컴퓨터교육"), "컴퓨터교육")
        self.assertEqual(engine._pick_major("컴퓨터공학"), "컴퓨터공학")

    def test_pick_major_keeps_abbreviation_fallback(self):
        engine = make_engine(["컴퓨터공학과", "생명공학과"])
        self.assertEqual(engine._pick_major("컴공"), "컴퓨터공학과")
        self.assertEqual(engine._pick_major("생명"), "생명공학과")

    def test_pick_major_prefers_shortest_substring_candidate(self):
        engine = make_engine(["의료경영학과", "경영학과", "글로벌경영학과"])
        self.assertEqual(engine._pick_major("경영학"), "경영학과")

    def test_pick_major_rejects_mid_substring_mismatch(self):
        # '마법학과'는 존재하지 않음 -> '법학과'(norm '법학'이 '마법학'의 중간일치)로
        # 오매칭되면 안 된다. 접두 매칭이므로 빈 문자열을 돌려준다.
        engine = make_engine(["법학과", "경영학과"])
        self.assertEqual(engine._pick_major("마법학과"), "")
        # 접두 매칭('경영학' -> '경영학과')은 여전히 유지되어야 한다.
        self.assertEqual(engine._pick_major("경영학"), "경영학과")

    def test_pick_school_matches_partial_name(self):
        engine = make_engine(schools=["신정고", "울산여고"])
        self.assertEqual(engine._pick_school("신정고 1학년입니다"), "신정고")
        self.assertEqual(engine._pick_school("없는학교"), "")


class TrackSubjectExpansionTest(unittest.TestCase):
    def test_expand_splits_compound_expression(self):
        engine = make_engine(master=["기하", "미적분Ⅱ"])
        subjects, groups, dropped = engine._expand_track_subject("기하 또는 미적분Ⅱ")
        self.assertEqual(subjects, ["기하", "미적분Ⅱ"])
        self.assertEqual(groups, [])
        self.assertEqual(dropped, [])

    def test_expand_classifies_group_terms(self):
        engine = make_engine(master=["물리학"])
        subjects, groups, dropped = engine._expand_track_subject("과학")
        self.assertEqual(subjects, [])
        self.assertEqual(groups, ["과학"])
        self.assertEqual(dropped, [])

    def test_expand_maps_legacy_misugbun(self):
        engine = make_engine(master=["미적분Ⅰ", "미적분Ⅱ"])
        subjects, groups, dropped = engine._expand_track_subject("미적분")
        self.assertEqual(subjects, ["미적분Ⅰ", "미적분Ⅱ"])
        self.assertEqual(dropped, [])

    def test_expand_drops_unmatched_subject(self):
        # 변경 1: 마스터 미매칭 & 비교과군 잔여는 과목으로 유지하지 않고 폐기한다
        engine = make_engine(master=["물리학"])
        subjects, groups, dropped = engine._expand_track_subject("로봇과 공학세계")
        self.assertEqual(subjects, [])
        self.assertEqual(groups, [])
        self.assertEqual(dropped, ["로봇과 공학세계"])


class NewGroupTermsTest(unittest.TestCase):
    def test_new_group_terms_are_classified_as_groups_not_dropped(self):
        engine = make_engine(master=["물리학"])
        terms = [
            "역사", "윤리", "지리", "일반사회", "과학교과", "수학1", "수학2",
            "한국사12", "전과목", "과학전과목", "수학전과목",
            "제2외국어관련과목", "제2외국어과목", "제3외국어과목",
        ]
        for term in terms:
            subjects, groups, dropped = engine._expand_track_subject(term)
            self.assertEqual(subjects, [], term)
            self.assertEqual(groups, [term], term)
            self.assertEqual(dropped, [], term)


class CircledDigitGroupTermTest(unittest.TestCase):
    def test_circled_digit_math_terms_are_groups_not_dropped(self):
        # 실측 대학트랙 원문은 로마숫자(Ⅰ/Ⅱ)가 아니라 원문자(①/②) 글리프를 사용
        engine = make_engine(master=["물리학"])
        for term in ("수학①", "수학②"):
            subjects, groups, dropped = engine._expand_track_subject(term)
            self.assertEqual(subjects, [])
            self.assertEqual(groups, [term])
            self.assertEqual(dropped, [])


class NewAliasExpansionTest(unittest.TestCase):
    def test_new_aliases_map_to_official_subjects(self):
        engine = make_engine(master=[
            "물리학", "한국지리 탐구", "동아시아 역사 기행", "미적분Ⅰ", "미적분Ⅱ",
            "프랑스어 회화", "도시의 미래 탐구", "과학과제 연구", "역학과 에너지",
            "영어 독해와 작문",
        ])
        cases = {
            "물리": ["물리학"],
            "한국지리": ["한국지리 탐구"],
            "동아시아사": ["동아시아 역사 기행"],
            "수학(특히 미적분)": ["미적분Ⅰ", "미적분Ⅱ"],
            "프랑스 회화": ["프랑스어 회화"],
            "도시와 미래탐구": ["도시의 미래 탐구"],
            "과학과제 탐구": ["과학과제 연구"],
            "물리과 에너지": ["역학과 에너지"],
            "독해와 작문": ["영어 독해와 작문"],
        }
        for raw, expected in cases.items():
            subjects, groups, dropped = engine._expand_track_subject(raw)
            self.assertEqual(subjects, expected, raw)
            self.assertEqual(dropped, [], raw)


class DroppedTermsLayer2Test(unittest.TestCase):
    def test_layer2_surfaces_dropped_terms_with_counts(self):
        engine = RecommendationEngine.__new__(RecommendationEngine)
        engine.store = DataStore(
            subject_master={norm("물리학"): "물리학"},
            uni_tracks=[
                UniTrackRow(university="A대", unit="컴퓨터공학과", subject="로봇과 공학세계",
                            track_type="권장", priority=1.0, comment=""),
                UniTrackRow(university="B대", unit="컴퓨터공학과", subject="로봇과 공학세계",
                            track_type="권장", priority=1.0, comment=""),
            ],
        )
        layer2 = engine._layer2_universities("컴퓨터공학과")
        self.assertNotIn("로봇과 공학세계", [s["subject"] for s in layer2["subjects"]])
        self.assertEqual(layer2["dropped_terms"], [{"term": "로봇과 공학세계", "count": 2}])


class OfferingChoiceFieldsLoadTest(unittest.TestCase):
    def test_load_curriculum_reads_choice_group_and_take_n(self):
        import openpyxl

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "고1편제표"
        ws.append(["학교", "학년도", "구분", "교과군", "과목유형", "과목명", "공식과목명", "매칭",
                   "기준학점", "운영학점", "1-1", "1-2", "2-1", "2-2", "3-1", "3-2",
                   "원본파일", "원본시트", "선택군ID", "택N"])
        ws.append(["테스트고", 2026, "학생선택", "수학", "일반", "확률과 통계", "확률과 통계", "O",
                   4, 4, None, None, 4, None, None, None,
                   "test.xlsx", "s", "테스트고-2-1-01", 3])
        ws.append(["테스트고", 2026, "학교지정", "정보", "공통", "정보", "정보", "O",
                   4, 4, 4, None, None, None, None, None,
                   "test.xlsx", "s", None, None])
        tmp = Path(__file__).resolve().parent / "_tmp_offering_load.xlsx"
        wb.save(tmp)
        try:
            engine = RecommendationEngine.__new__(RecommendationEngine)
            engine.store = DataStore()
            engine.curriculum_path = tmp
            engine._load_curriculum(engine.store)
            offerings = {o.subject: o for o in engine.store.schools["테스트고"]}
            self.assertEqual(offerings["확률과 통계"].choice_group, "테스트고-2-1-01")
            self.assertEqual(offerings["확률과 통계"].take_n, 3)
            self.assertEqual(offerings["정보"].choice_group, "")
            self.assertEqual(offerings["정보"].take_n, 0)
            self.assertIn(norm("확률과 통계"), engine.store.all_offered_keys)
            self.assertIn(norm("정보"), engine.store.all_offered_keys)
        finally:
            tmp.unlink(missing_ok=True)


class Layer3ModeTest(unittest.TestCase):
    def test_layer3_marks_designated_vs_choice_group_mode(self):
        engine = make_engine()
        engine.store.schools["신정고"] = [
            Offering(subject="확률과 통계", matched=True, section="학생선택", group="수학",
                     subject_type="일반", semesters=("2-1",),
                     choice_group="신정고-2-1-01", take_n=3),
            Offering(subject="물리학", matched=True, section="학교지정", group="과학",
                     subject_type="공통", semesters=("1-1",)),
        ]
        engine.store.all_offered_keys = {norm("확률과 통계"), norm("물리학")}
        layer1 = {"categories": []}
        layer2 = {
            "subjects": [
                {"subject": "확률과 통계", "type": "핵심", "university_count": 1, "rate": None},
                {"subject": "물리학", "type": "핵심", "university_count": 1, "rate": None},
            ],
            "group_mentions": [],
        }
        layer3 = engine._layer3_school("신정고", layer1, layer2)
        by_subject = {a["subject"]: a for a in layer3["available"]}
        self.assertEqual(by_subject["확률과 통계"]["mode"], "선택군")
        self.assertEqual(by_subject["확률과 통계"]["take_n"], 3)
        self.assertEqual(by_subject["확률과 통계"]["choice_group"], "신정고-2-1-01")
        self.assertEqual(by_subject["물리학"]["mode"], "지정")

    def test_layer3_designation_wins_when_subject_has_both_plain_and_group_rows(self):
        engine = make_engine()
        engine.store.schools["신정고"] = [
            Offering(subject="물리학", matched=True, section="학교지정", group="과학",
                     subject_type="공통", semesters=("1-1",)),
            Offering(subject="물리학", matched=True, section="학생선택", group="과학",
                     subject_type="일반", semesters=("2-1",),
                     choice_group="신정고-2-1-02", take_n=2),
        ]
        engine.store.all_offered_keys = {norm("물리학")}
        layer1 = {"categories": []}
        layer2 = {
            "subjects": [{"subject": "물리학", "type": "핵심", "university_count": 1, "rate": None}],
            "group_mentions": [],
        }
        layer3 = engine._layer3_school("신정고", layer1, layer2)
        self.assertEqual(layer3["available"][0]["mode"], "지정")


class CommonMissingScopeTest(unittest.TestCase):
    def test_unavailable_scope_marks_common_missing_only_when_offered_nowhere(self):
        engine = make_engine()
        engine.store.schools = {
            "신정고": [Offering(subject="물리학", matched=True, section="학생선택", group="과학",
                               subject_type="일반", semesters=("2-1",))],
            "울산여고": [Offering(subject="화학", matched=True, section="학생선택", group="과학",
                                subject_type="일반", semesters=("2-1",))],
        }
        engine.store.all_offered_keys = {norm("물리학"), norm("화학")}
        layer1 = {"categories": []}
        layer2 = {
            "subjects": [
                {"subject": "화학", "type": "권장", "university_count": 1, "rate": None},
                {"subject": "문예창작", "type": "권장", "university_count": 1, "rate": None},
            ],
            "group_mentions": [],
        }
        layer3 = engine._layer3_school("신정고", layer1, layer2)
        missing = {m["subject"]: m for m in layer3["unavailable"]}
        # 화학: 신정고엔 없지만 울산여고엔 있음 -> 공통 미개설 아님
        self.assertEqual(missing["화학"]["scope"], "")
        # 문예창작: 로드된 어느 학교에도 없음 -> 공통 미개설
        self.assertEqual(missing["문예창작"]["scope"], "공통 미개설")


if __name__ == "__main__":
    unittest.main()
