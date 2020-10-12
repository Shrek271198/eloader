from eload import (FieldEquipment, LightingEquipment, MechanicalEquipment,
                   MotorControlCenter, UPSEquipment, client_mel_builder,
                   mcc_builder, me_builder, read_standards, read_client_mel, ElectricalLoadSummary)
from utility import round_up


def test_me_1():
    row = (121, 'CN', '001', 'PRIMARY CRUSHER JIB CRANE ', 2, 'MCC-001', 20, 'DOL', 415, 'DUTY', 'A', 5)
    me = me_builder(row)

    assert me.installed_kw == 20
    assert me.power_factor == 0.83
    assert me.kva == 25.6
    assert me.load_factor == 0.5
    assert me.diversity_utilisation == 0.1
    assert me.avg_load_factor == 0.05
    assert me.max_kw == 10
    assert me.max_kvar == 6.7
    assert me.max_kva == 12.8
    assert me.avg_load_kw == 1
    assert me.spare_capacity == 0.26


def test_me_2():
    row = (121, 'CP', '001', 'PRIMARY CRUSHING AIR COMPRESSOR ', 2, 'MCC-001', 7.5, 'FEEDER', 415, 'DUTY', 'A', 5)
    me = me_builder(row)
    assert me.installed_kw == 7.5
    assert me.power_factor == 0.85
    assert me.kva == 9.6
    assert me.load_factor == 0.9
    assert me.diversity_utilisation == 0.95
    assert me.avg_load_factor == 0.855
    assert me.max_kw == 6.8
    assert me.max_kvar == 4.2
    assert me.max_kva == 8.6
    assert me.avg_load_kw == 6.4
    assert me.spare_capacity == 1.64


def test_me_3():
    row = (121, 'CR', '001', 'PRIMARY CRUSHER ', 2, 'MCC-001', 10, 'VSD', 415, 'DUTY', 'A', 5)
    me = me_builder(row)

    assert me.tag_number =='121CR001'
    assert me.installed_kw == 10
    assert me.power_factor == 0.9
    assert me.kva == 12.1
    assert me.load_factor == 0.8
    assert me.diversity_utilisation == 0.6
    assert me.avg_load_factor == 0.48
    assert me.max_kw == 8
    assert me.max_kvar == 3.8
    assert me.max_kva == 9.7
    assert me.avg_load_kw == 4.8
    assert me.spare_capacity == 1.16


def test_me_4():
    row = (121, 'CV', '005', 'PRIMARY FEEDER DRIBBLE CONVEYOR ', 2, 'MCC-001', 2.2, 'DOL', 415, 'DUTY', 'A', 5)
    me = me_builder(row)

    assert me.installed_kw == 2.2
    assert me.power_factor == 0.81
    assert me.kva == 3.1
    assert me.load_factor == 0.9
    assert me.diversity_utilisation == 0.95
    assert me.avg_load_factor == 0.855
    assert me.max_kw == 2
    assert me.max_kvar == 1.4
    assert me.max_kva == 2.8
    assert me.avg_load_kw == 1.9
    assert me.spare_capacity == 0.54


def test_me_5():
    row = (121, 'DR', '001', 'PRIMARY CRUSHING AIR COMPRESSED AIR DRYER ', 2, 'MCC-001', 0.1, 'DOL', 415, 'DUTY', 'A', 5)
    me = me_builder(row)

    assert me.tag_number =='121DR001'
    assert me.installed_kw == 0.1
    assert me.power_factor == 0.55
    assert me.kva == 0.3
    assert me.load_factor == 0.8
    assert me.diversity_utilisation == 0.95
    assert me.avg_load_factor == 0.76
    assert me.max_kw == 0.1
    assert me.max_kvar == 0.2
    assert me.max_kva == 0.2
    assert me.avg_load_kw == 0.1
    assert me.spare_capacity == 0.04


def test_mcc():
    lighting_load = 0
    ups_load = 3
    fe_dist_load = 45

    rows = []
    rows.append((121, 'CN', '001', 'PRIMARY CRUSHER JIB CRANE ', 2, 'MCC-001', 20, 'DOL', 415, 'DUTY', 'A', 5))
    rows.append((121, 'CP', '001', 'PRIMARY CRUSHING AIR COMPRESSOR ', 2, 'MCC-001', 7.5, 'FEEDER', 415, 'DUTY', 'A', 5))
    rows.append((121, 'CR', '001', 'PRIMARY CRUSHER ', 2, 'MCC-001', 10, 'VSD', 415, 'DUTY', 'A', 5))
    rows.append((121, 'CV', '005', 'PRIMARY FEEDER DRIBBLE CONVEYOR ', 2, 'MCC-001', 2.2, 'DOL', 415, 'DUTY', 'A', 5))
    rows.append((121, 'DR', '001', 'PRIMARY CRUSHING AIR COMPRESSED AIR DRYER ', 2, 'MCC-001', 0.1, 'DOL', 415, 'DUTY', 'A', 5))

    mcc_me_list = []
    for row in rows:
        me = me_builder(row)
        mcc_me_list.append(me)

    mcc = mcc_builder("mcc-001", lighting_load, ups_load, fe_dist_load, mcc_me_list)

    assert mcc.total_installed_kw == 87.8
    assert mcc.total_kva == 98.7
    assert mcc.total_max_kw == 63.0
    assert mcc.total_max_kvar == 52.4
    assert mcc.total_max_kva == 70.2
    assert mcc.total_avg_load_kw == 46.6

    assert mcc.spare_starters == 2.0
    assert mcc.contingency_load == 11.0
    assert mcc.total_spare_allocation == 22.0
    assert mcc.total_mcc_load_allowed == 109.8

    assert mcc.max_voltage == 415
    assert mcc.contingency_factor == 4.0
    assert mcc.total_actual_contingency == 74.20
    assert mcc.tx_size == 500
    assert mcc.spare_tx == 85


def test_els():

    standards_file= "tests/fixtures/standards.xlsx"
    mel_file= "tests/fixtures/mel.xlsx"

    # Read in Project Standards excel file
    STANDARD = read_standards(standards_file)

    # Read in client Mechanical Equipment List
    MEL = read_client_mel(mel_file)

    # Initialise Motor Control Centers
    MCC_List = []
    for number in MEL.mcc_numbers:

        # Filter equipment by mcc_number
        mcc_mel = []
        for me in MEL.mel:
            if me.mcc_number == number:
                mcc_mel.append(me)

        MCC = mcc_builder(
            number,
            STANDARD.lighting_load,
            STANDARD.ups_load,
            STANDARD.fe_dist_load,
            mcc_mel,
        )

        MCC_List.append(MCC)

    els = ElectricalLoadSummary(MCC_List)

    assert els.connected_load_kw == 546
    assert els.connected_load_kva == 642
    assert els.max_demand_kw == 413
    #assert els.max_demand_kvar == 298
#    max_demand_kva
#    ave_load_kva
#    contingency_factor_kva
#    total_actual_contingency
#    tx_size
#    spare_tx
#
#    network_loss_kw
#    network_loss_kvar
#    network_loss_kva
