import unittest

from roadmap_rag.graph_rag import GraphRAGEngine, RoadmapGraph


class MatchingTest(unittest.TestCase):
    def make_engine(self, majors):
        engine = GraphRAGEngine.__new__(GraphRAGEngine)
        engine.graph = RoadmapGraph(majors={major: [] for major in majors})
        return engine

    def test_pick_major_prefers_existing_major_before_abbreviation_replacement(self):
        engine = self.make_engine(["생명과학", "생명공학", "컴퓨터교육", "컴퓨터공학"])

        self.assertEqual(engine._pick_major("생명과학"), "생명과학")
        self.assertEqual(engine._pick_major("컴퓨터교육"), "컴퓨터교육")
        self.assertEqual(engine._pick_major("컴퓨터공학"), "컴퓨터공학")

    def test_pick_major_keeps_abbreviation_fallback(self):
        engine = self.make_engine(["컴퓨터공학", "생명공학"])

        self.assertEqual(engine._pick_major("컴공"), "컴퓨터공학")
        self.assertEqual(engine._pick_major("생명"), "생명공학")

    def test_match_graph_major_exact_base_key(self):
        engine = self.make_engine([])
        graph = RoadmapGraph(majors={"건축학": []})

        self.assertEqual(engine._match_graph_major("건축학과", graph), "건축학")

    def test_match_graph_major_does_not_use_one_character_substring_fallback(self):
        engine = self.make_engine([])
        graph = RoadmapGraph(majors={"화학과": []})

        self.assertEqual(engine._match_graph_major("과", graph), "과")


if __name__ == "__main__":
    unittest.main()
