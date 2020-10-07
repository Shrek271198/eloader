import math
from datetime import date
from enum import Enum
from typing import List

from data import LOAD_FACTOR, MOTOR_POWER_FACTOR
from dataclasses import dataclass, field
from utility import round_up as round
from utility import vlookup


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

    lighting_load: float
    ups_load: float
    fe_dist_load: float

    min_thermistor: float
    max_thermistor: float

    min_rtd: float
    max_rtd: float

    substation_cost: float
    max_cable_size: float


@dataclass
class LightingEquipment:
    """Class for Lighting Equipment details."""

    installed_kw: float

    power_factor: float = 1
    load_factor: float = 0.75
    diversity_utilisation: float = 0.9
    efficiency: float = 1

    kva: float = field(init = False)
    avg_load_factor: float = field(init = False)
    max_kw: float = field(init = False)
    max_kvar: float = field(init = False)
    max_kva: float = field(init = False)
    avg_load_kw: float = field(init = False)

    def __post_init__(self):

        # Init kva
        try:
            self.kva = round(self.installed_kw/self.efficiency/self.power_factor, 1)
        except:
            self.kva = 0

        # Init avg_load_factor
        self.avg_load_factor = round(self.load_factor*self.diversity_utilisation, 3)

        # Init max_kw
        try:
            self.max_kw = round(self.installed_kw*self.load_factor, 1)
        except:
            self.max_kw = 0

        # Init max_kvar
        self.max_kvar = self.max_kw

        # Init max_kva
        self.max_kva = round(self.kva*self.load_factor, 1)

        # Init max_load_kw
        try:
            self.avg_load_kw = round(self.installed_kw*self.avg_load_factor, 1)
        except:
            self.avg_load_kw = 0


@dataclass
class UPSEquipment:
    """Class for UPS Equipment details."""

    installed_kw: float

    power_factor: float = 1
    load_factor: float = 0.75
    diversity_utilisation: float = 0.9
    efficiency: float = 1

    kva: float = field(init = False)
    avg_load_factor: float = field(init = False)
    max_kw: float = field(init = False)
    max_kvar: float = field(init = False)
    max_kva: float = field(init = False)
    avg_load_kw: float = field(init = False)

    def __post_init__(self):

        # Init kva
        try:
            self.kva = round(self.installed_kw/self.efficiency/self.power_factor, 1)
        except:
            self.kva = 0

        # Init avg_load_factor
        self.avg_load_factor = round(self.load_factor*self.diversity_utilisation, 3)

        # Init max_kw
        try:
            self.max_kw = round(self.installed_kw*self.load_factor, 1)
        except:
            self.max_kw = 0

        # Init max_kvar
        self.max_kvar = self.max_kw

        # Init max_kva
        self.max_kva = round(self.kva*self.load_factor, 1)

        # Init max_load_kw
        try:
            self.avg_load_kw = round(self.installed_kw*self.avg_load_factor, 1)
        except:
            self.avg_load_kw = 0


@dataclass
class FieldEquipment:
    """Class for Field Equipment details."""

    installed_kw: float

    power_factor: float = 1
    load_factor: float = 0.75
    diversity_utilisation: float = 0.9
    efficiency: float = 1

    kva: float = field(init = False)
    avg_load_factor: float = field(init = False)
    max_kw: float = field(init = False)
    max_kvar: float = field(init = False)
    max_kva: float = field(init = False)
    avg_load_kw: float = field(init = False)

    def __post_init__(self):

        # Init kva
        try:
            self.kva = round(self.installed_kw/self.efficiency/self.power_factor, 1)
        except:
            self.kva = 0

        # Init avg_load_factor
        self.avg_load_factor = round(self.load_factor*self.diversity_utilisation, 3)

        # Init max_kw
        try:
            self.max_kw = round(self.installed_kw*self.load_factor, 1)
        except:
            self.max_kw = 0

        # Init max_kvar
        self.max_kvar = self.max_kw

        # Init max_kva
        self.max_kva = round(self.kva*self.load_factor, 1)

        # Init max_load_kw
        try:
            self.avg_load_kw = round(self.installed_kw*self.avg_load_factor, 1)
        except:
            self.avg_load_kw = 0


@dataclass
class MechanicalEquipment:
    """Class for Mechanical Equipment details."""

    area: str
    type: str
    number: str
    name: str
    workpack: str
    mcc_number: str
    installed_kw: float
    starter_type: str
    voltage: float
    operation_mode: OperationMode
    rev: str
    procurement_rating: int
    tag_number: str = field(init = False)

    power_factor: float = field(init = False)
    efficiency: float = field(init = False)
    kva: float = field(init = False)
    load_factor: float = field(init = False)
    diversity_utilisation: float = field(init = False)
    avg_load_factor: float = field(init = False)
    max_kw: float = field(init = False)
    max_kvar: float = field(init = False)
    max_kva: float = field(init = False)
    avg_load_kw: float = field(init = False)

    def __post_init__(self):

        # Init tag_number
        self.tag_number = self.area + self.type + self.number

        # Init efficiency
        self.efficiency = vlookup(self.installed_kw, MOTOR_POWER_FACTOR, 1)

        # Init power_factor
        if self.starter_type in ['VSD', 'VSD Dual']:
            self.power_factor = 0.9
        else:
            self.power_factor = vlookup(self.installed_kw, MOTOR_POWER_FACTOR, 2)

        # Init kva
        self.kva = round(self.installed_kw/self.efficiency/self.power_factor, 1)

        # Init load_factor
        if self.operation_mode == 2:
            self.load_factor = 0
        else:
            self.load_factor = vlookup(self.type, LOAD_FACTOR, 2)

        # Init diversity_utilisation
        if self.operation_mode == 2:
            self.diversity_utilisation = 0
        else:
            self.diversity_utilisation = vlookup(self.type, LOAD_FACTOR, 3)

        # Init avg_load_factor
        self.avg_load_factor = round(self.load_factor*self.diversity_utilisation, 2)

        # Init max_kw
        self.max_kw = round(self.installed_kw*self.load_factor, 1)

        # Init max_kvar
        self.max_kvar = round(self.max_kw*math.tan(math.acos(self.power_factor)), 1)

        # Init max_kva
        self.max_kva = round(self.kva*self.load_factor, 1)

        # Init max_load_kw
        self.avg_load_kw = round(self.installed_kw*self.avg_load_factor, 1)


@dataclass
class MotorControlCenter:
    """Class for storing the electrical load details of a MCC."""

    lighting: LightingEquipment
    ups: UPSEquipment
    field_equipment: FieldEquipment

    mel: List[MechanicalEquipment] = field(default_factory=list)

    ##total_installed_kw
    ##total_kva
    ##total_max_kw
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

    # Init MCCs
    MCC = {}
    for number in MEL.mcc_numbers:
        MCC[number] = MotorControlCenter(LightingEquipment(STANDARD.lighting_load),
                                         UPSEquipment(STANDARD.ups_load),
                                         FieldEquipment(STANDARD.fe_dist_load))

    for me in MEL.mel:
        MCC[me.mcc_number].mel.append(me)
