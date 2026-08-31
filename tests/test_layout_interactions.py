import unittest

from app import app


class LayoutInteractionTests(unittest.TestCase):
    def setUp(self):
        app.config.update(TESTING=True)
        self.client = app.test_client()

    def test_table_seat_reduction_preserves_occupied_seats_before_eviction(self):
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("const staleOccupiedKeys", html)
        self.assertIn("const availableKeys", html)
        self.assertIn("座位数量减少后容量不足", html)
        self.assertIn("syncSeatKeys(desiredKeys, true)", html)

    def test_table_size_changes_reflow_party_order_stably(self):
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("collectRoleOccupants", html)
        self.assertIn("placeReflowOccupants", html)
        self.assertIn("hasTopologyChanged", html)
        self.assertIn("sortRoundtableRoleKeys", html)
        self.assertIn("if(shouldReflow) clearReflowOccupants", html)
        self.assertIn("buildRoundtablePartyPlan(availableRoundtableKeys", html)
        self.assertIn("layout.longtableConfig.partySeatRule = assignConfig.rule", html)

    def test_longtable_supports_shared_two_sided_overflow_for_both_parties(self):
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("layout.longtableConfig.guestSide", html)
        self.assertIn("setLongtablePartySide('host'", html)
        self.assertIn("setLongtablePartySide('guest'", html)
        self.assertIn("sortLongtableExtensionSeatKeys", html)
        self.assertIn("buildLongtablePartyPlan", html)
        self.assertIn("key.startsWith('lt-left-') || key.startsWith('lt-right-')", html)
        self.assertIn("primarySide === 'bottom' ? b.pos - a.pos : a.pos - b.pos", html)
        self.assertIn("if(sideLists.left[index]) ordered.push", html)
        self.assertIn("if(sideLists.right[index]) ordered.push", html)
        self.assertIn("guestExtensionKeys", html)

    def test_classroom_podium_and_aisle_numbering_helpers_are_exposed(self):
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("getPodiumColumnOrder", html)
        self.assertIn("centerLeft + offset", html)
        self.assertIn("centerLeft - offset", html)
        self.assertIn("renumberMainColumnsAfterAisles", html)
        self.assertIn("isFullMainAisleColumn", html)

    def test_roundtable_dual_principal_keeps_seat_keys_and_rotates_visual_axis(self):
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("layout.roundtableConfig.dualPrincipal", html)
        self.assertIn("const principalOffset", html)
        self.assertIn("step * pos - Math.PI / 2 - principalOffset", html)
        self.assertIn("主人1号与宾客1号", html)

    def test_u_table_uses_stable_leader_near_order_and_reflow(self):
        html = self.client.get("/").get_data(as_text=True)
        self.assertIn("const sortUtableSeatKeys", html)
        self.assertIn("getPodiumColumnOrder(leaderCount)", html)
        self.assertIn("layout.utableConfig.leftColumnCount", html)
        self.assertIn("layout.utableConfig.rightColumnCount", html)
        self.assertIn("for(let column = 0; column <= maxSideColumn; column++)", html)
        self.assertIn("if(leftKey) ordered.push(leftKey)", html)
        self.assertIn("if(rightKey) ordered.push(rightKey)", html)
        self.assertIn("flexDirection: layout.utableConfig.leaderSide === 'bottom' ? 'row-reverse' : 'row'", html)
        self.assertIn("const attendeeRank = new Map(attendees.value.map", html)
        self.assertIn("placeOrderedOccupants(occupants, availableKeys)", html)
        self.assertIn('@input="scheduleUtableSeatSync"', html)


if __name__ == "__main__":
    unittest.main()
