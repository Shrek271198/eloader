from datetime import date
from enum import Enum
from typing import List

from dataclasses import dataclass, field


@dataclass
class ProjectStandard:
    """ Class for storing Project standard details."""
    field_isolator: bool
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

    substation_cost: float
    max_cable_size: float

class OperationMode(Enum):
    DUTY = 1
    STANDBY = 2

@dataclass
class MechanicalEquipment:
    """Class for Mechanical Equipment details."""
    __slots__ = ['area', 'type', 'number','name', 'workpack', 'mcc_number', 'installed_power', 'starter_type', 'voltage','operation_mode','rev','procurement_rating','rating_description']
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
    rating_description: str

    #installed
    #pf
    #kva
    #load_factor
    #diversity_utilisation
    #avg_load_factor
    #max_kw
    #max_kvar
    #avg_load_kw

    def tag_number(self):
        return self.area+self.type+self.number


@dataclass
class MotorControlCenter:
    """Class for storing the electrical load details of a MCC."""

    mel: List[MechanicalEquipment] = field(default_factory=list)

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

    #total_ave_load
    #substation_load_dist
    #ss_ladders
    #contingency_factor
    #spare_starters
    #avg_starter_load
    #total_spare_allocation
    #total_mcc_load_allowed

    #total_contingency_factor
    #total

    #tx_size
    #spare_tx


@dataclass
class MCCLoadList:
    """Class for MCC Load summaries"""

    load: List[MotorControlCenter]

    #total_connected_load_kw
    #total_connected_load_kva
    #max_demand_kw
    #max_demand_kvar
    #max_demand_kva
    #total_ave_load
    #total_contingency_factor
    #total


def eload(standards, mel):

    print(standards)
    print(mel)
