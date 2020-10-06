from datetime import date
from enum import Enum
from typing import List

from dataclasses import dataclass, field


class OperationMode(Enum):
    DUTY = 1
    STANDBY = 2


@dataclass
class ProjectStandard:
    """ Class for storing Project standard details."""

    field_isolator: str
    project_name: str
    folder_name: str
    folder_path: str

    mcc_to_vsd: int
    mcc_to_fi: int
    vsd_to_fi: int
    vsd_mcc_to_motor: int

    project: str
    revision: str
    prepared_by: str
    date_prepared: date
    approved_by: str
    date_approved: date

    avg_count: float
    starters: float

    misc_load: float
    ups_load: float
    fe_dist_load: float

    min_thermistor: float
    max_thermistor: float

    min_rtd: float
    max_rtd: float

    substation_cost: float
    max_cable_size: float


@dataclass
class MechanicalEquipment:
    """Class for Mechanical Equipment details."""

    area: str
    type: str
    number: str
    name: str
    workpack: str
    mcc_number: str
    installed_power: float
    starter_type: str
    voltage: float
    operation_mode: OperationMode
    rev: str
    procurement_rating: int
    tag_number: str = field(init = False)

    # installed
    # pf
    # kva
    # load_factor
    # diversity_utilisation
    # avg_load_factor
    # max_kw
    # max_kvar
    # avg_load_kw

    def __post_init__(self):

        # Init tag_number
        self.tag_number = self.area + self.type + self.number


@dataclass
class MotorControlCenter:
    """Class for storing the electrical load details of a MCC."""

    list: List[MechanicalEquipment] = field(default_factory=list)

    ##total_installed_kw
    ##total_kva
    ##total_max_kvar
    ##total_max_kva
    ##total_avg_load_kw

    ##total_connected_load_kw
    ##total_connected_load_kva
    ##max_demand_kw
    ##max_demand_kvar
    ##max_demand_kva

    # total_ave_load
    # substation_load_dist
    # ss_ladders
    # contingency_factor
    # spare_starters
    # avg_starter_load
    # total_spare_allocation
    # total_mcc_load_allowed

    # total_contingency_factor
    # total

    # tx_size
    # spare_tx


@dataclass
class MechanicalEquipmentList:
    """Class for Mechanical Equipment List."""

    mel: List[MechanicalEquipment] = field(default_factory=list)
    mcc_numbers: list = field(init = False)
    mcc: dict = field(init = False)

    def __post_init__(self):

        # Init mcc_numbers
        self.mcc_numbers = []
        for me in self.mel:
            if me.mcc_number not in self.mcc_numbers:
                self.mcc_numbers.append(me.mcc_number)


        # Init mcc
        self.mcc = {}
        for number in self.mcc_numbers:
            self.mcc[number] = MotorControlCenter()

        for me in self.mel:
            self.mcc[me.mcc_number].list.append(me)





@dataclass
class MCCLoadList:
    """Class for MCC Load summaries"""

    load: List[MotorControlCenter]

    # total_connected_load_kw
    # total_connected_load_kva
    # max_demand_kw
    # max_demand_kvar
    # max_demand_kva
    # total_ave_load
    # total_contingency_factor
    # total



def read_standards(standards):
    """ Create ProjectStandard object from a Project Standards excel file."""

    from openpyxl import load_workbook

    wb = load_workbook(filename=standards, data_only=True)
    ws = wb.active

    standards = ProjectStandard(
        ws["E5"].value,
        ws["E10"].value,
        ws["E11"].value,
        ws["E12"].value,
        ws["E15"].value,
        ws["E16"].value,
        ws["E17"].value,
        ws["E18"].value,
        ws["E21"].value,
        ws["E22"].value,
        ws["E23"].value,
        ws["E24"].value,
        ws["E25"].value,
        ws["E26"].value,
        0,
        0,
        ws["E37"].value,
        ws["E38"].value,
        ws["E39"].value,
        ws["C43"].value,
        ws["E43"].value,
        ws["C44"].value,
        ws["E44"].value,
        ws["D47"].value,
        ws["D49"].value,
    )

    return standards


def read_mel(mel):
    """Create a list of MechanicalEquipment objects from a MEL excel file."""

    from openpyxl import load_workbook

    wb = load_workbook(filename=mel,  data_only=True)
    ws = wb.active

    mel = []

    for row in ws.iter_rows(min_row=8, max_col=12, max_row=ws.max_row, values_only=True):
        me = MechanicalEquipment(
            str(row[0]),
            str(row[1]),
            str(row[2]),
            str(row[3]),
            str(row[4]),
            str(row[5]),
            float(row[6]),
            str(row[7]),
            float(row[8]),
            row[9],
            str(row[10]),
            int(row[11]),
        )

        mel.append(me)

    MEL = MechanicalEquipmentList(mel)

    return MEL


def eload(standards, mel):

    STANDARD = read_standards(standards)
    MEL = read_mel(mel)

    import pdb
    pdb.set_trace()




