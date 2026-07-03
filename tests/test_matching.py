# -*- coding: utf-8 -*-
import unittest

from roadmap_rag.graph_rag import DataStore, GraphRAGEngine, norm


def make_engine(depts=(), master=(), schools=()):
    engine = GraphRAGEngine.__new__(GraphRAGEngine)
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

    def test_pick_school_matches_partial_name(self):
        engine = make_engine(schools=["신정고", "울산여고"])
        self.assertEqual(engine._pick_school("신정고 1학년입니다"), "신정고")
        self.assertEqual(engine._pick_school("없는학교"), "")


class TrackSubjectExpansionTest(unittest.TestCase):
    def test_expand_splits_compound_expression(self):
        engine = make_engine(master=["기하", "미적분Ⅱ"])
        subjects, groups = engine._expand_track_subject("기하 또는 미적분Ⅱ")
        self.assertEqual(subjects, ["기하", "미적분Ⅱ"])
        self.assertEqual(groups, [])

    def test_expand_classifies_group_terms(self):
        engine = make_engine(master=["물리학"])
        subjects, groups = engine._expand_track_subject("과학")
        self.assertEqual(subjects, [])
        self.assertEqual(groups, ["과학"])

    def test_expand_maps_legacy_misugbun(self):
        engine = make_engine(master=["미적분Ⅰ", "미적분Ⅱ"])
        subjects, groups = engine._expand_track_subject("미적분")
        self.assertEqual(subjects, ["미적분Ⅰ", "미적분Ⅱ"])

    def test_expand_keeps_unknown_subject_as_is(self):
        engine = make_engine(master=["물리학"])
        subjects, groups = engine._expand_track_subject("로봇과 공학세계")
        self.assertEqual(subjects, ["로봇과 공학세계"])
        self.assertEqual(groups, [])


if __name__ == "__main__":
    unittest.main()
