import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from app import unit_master


def test_dongs_and_floors_are_loaded_sorted():
    dongs = unit_master.get_dongs("봉담자이 프라이드시티")
    assert dongs
    assert dongs[0].endswith("동")

    floors = unit_master.get_floors("봉담자이 프라이드시티", dongs[0])
    assert floors == sorted(floors)


def test_unit_info_and_total_floor_are_returned():
    dong = unit_master.get_dongs("힐스테이트봉담프라이드시티")[0]
    floor = unit_master.get_floors("힐스테이트봉담프라이드시티", dong)[0]
    ho = unit_master.get_hos("힐스테이트봉담프라이드시티", dong, floor)[0]

    info = unit_master.get_unit_info("힐스테이트봉담프라이드시티", dong, floor, ho)
    assert info["type"]
    assert info["supply_m2"] > 0
    assert info["pyeong"] > 0

    total_floor = unit_master.get_total_floor("힐스테이트봉담프라이드시티", dong)
    assert total_floor >= floor
