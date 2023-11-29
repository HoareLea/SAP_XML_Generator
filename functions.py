"""Holds the main functions to process and outputs the XML file"""

import xml.etree.ElementTree as ET
from xml.dom import minidom
import math
import pandas as pd
from TBs import TBs
from levels_naming import levels_naming
import uuid
import datetime
import os

class ErrorFound(Exception):
    """Basic Class to raise exception"""
    pass

def generate_unique_name(base_name):
    """Generate a unique name for the temp XML files to be saved locally"""
    timestamp = datetime.datetime.now()
    unique_string = str(uuid.uuid4().hex)
    run_name = f"{base_name}_{str(timestamp).replace(':','-')}_{unique_string}"
    return run_name

def delete_files_in_directory(directory):
    """Deleting all temp files"""
    # Get the list of files in the directory
    file_list = os.listdir(directory)

    # Iterate over each file and delete it
    for file_name in file_list:
        file_path = os.path.join(directory, file_name)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)
                print(f"Deleted: {file_path}")
        except Exception as e:
            print(f"Error deleting {file_path}: {e}")
    pass

def data_to_xml(dictionary, root_name='AssessmentFull'):
    """Converts the dictionary element to a string suitable for XML"""
    root = ET.Element(root_name)
    def _data_to_xml(element, data):
        if isinstance(data, dict):
            for key, value in data.items():
                if isinstance(value, (list, dict)):
                    sub_element = ET.SubElement(element, key)
                    _data_to_xml(sub_element, value)
                else:
                    ET.SubElement(element, key).text = str(value)
        elif isinstance(data, list):
            for item in data:
                sub_element = ET.SubElement(element, 'item')
                _data_to_xml(sub_element, item)
        else:
            element.text = str(data)
    _data_to_xml(root, dictionary)
    return ET.tostring(root, encoding='unicode',xml_declaration=None)

def prettify(elem):
    """Formats the outputs with the releveant XML criteria"""
    reparsed = minidom.parseString(elem)
    prettified = reparsed.toprettyxml(indent="  ")
    final = prettified.replace('<?xml version="1.0" ?>\n', '')
    return final

def check_missing_data(unit, sheet, input_list):
    """Checks for any missing data and raises error if there's a mismatch"""
    for info in input_list:
        non_nan = sheet[info][sheet[info].notna()]
        if len(non_nan)==0:
            raise ErrorFound(f'Error for {unit["propertyName"]}: "{info}" has not been entered')
        unit[info] = sheet[info].tolist()[1]

def check_opening_type_data(unit, sheet, main, inputs, this_type):
    """Counts Openings on input types and raises error if there's a mismatch"""
    sheet = sheet[sheet['Type']==this_type]
    filtered_main = sheet[main][sheet[main].notna()]
    for data in inputs:
        filtered_input = sheet[data][sheet[data].notna()]
        if len(filtered_input)!=len(filtered_main):
            raise ErrorFound(f'''Error for {unit["propertyName"]}:
                             There is a mismatch between the number of "{main}" 
                             elements and the corresponding "{data}"''')

def check_openings_data(unit, sheet, main, inputs):
    """Counts Openings on input and raises error if there's a mismatch"""
    filtered_main = sheet[main][sheet[main].notna()]
    for data in inputs:
        filtered_input = sheet[data][sheet[data].notna()]
        if len(filtered_input)!=len(filtered_main):
            raise ErrorFound(f'''Error for {unit["propertyName"]}:
                             There is a mismatch between the number of "{main}" 
                             elements and the corresponding "{data}"''')

def check_op_element_names(unit, sheet):
    """Checks Element Names to ensure each is unique"""
    if len(set(sheet['Element name'])) != len(sheet['Element name']):
        raise ErrorFound(f'''Error for {unit["propertyName"]}:
                         Two or more opaque element entries have the same name. 
                         All elements require a unique name.''')

def input_reader(sheet):
    """Takes in the excel sheet to begin transforming into a dictionary"""
    # Instantiate unit dict. this will hold all the inputs from the excel sheet
    unit = {}

    # General information
    unit['propertyName'] = sheet['Property name'].tolist()[1]
    gen_info_list = [
        'Dwelling orientation',
        'Calculation type',
        'Terrain type',
        'Property type 1',
        'Property type 2',
        'Position of flat',
        'Which floor',
        'Tot no. storeys in block',
        'No. storeys',
        'Date built',
        'Sheltered sides',
        'Sunlight/sunshade',
        'Thermal mass parameter',
        'Living area'
    ]
    check_missing_data(unit, sheet, gen_info_list)

    # Level information
    # unit['floorToSlab'] = sheet['Floor to slab'].tolist()[1:]
    unit['floorToSlab'] = sheet['Floor to slab'][sheet['Floor to slab'].notna()].tolist()
    # unit['heatedIntArea'] = sheet['Heated internal floor area'].tolist()[1:]
    unit['heatedIntArea'] = sheet['Heated internal floor area'][sheet['Heated internal floor area'].notna()].tolist()
    # unit['heatLossPerim'] = sheet['Heat loss perimeter'].tolist()[1:]
    unit['heatLossPerim'] = sheet['Heat loss perimeter'][sheet['Heat loss perimeter'].notna()].tolist()
    
    # Opaque elements
    unit['opaqElementLevel'] = sheet['Level of opaque element'].tolist()[1:]
    # unit['opaqElementLevel'] = sheet['Level of opaque element'][sheet['Level of opaque element'].notna()].tolist() # THIS IS NOT APPLICABLE AT THIS STAGE BECAUSE THE MATCH_XML FUNCTION REQUIRES A DATAFRAME WITH ALL ELEMENT PROPERTIES WITH SAME LENGTH
    unit['opaqElementType'] = sheet['Element type'].tolist()[1:]
    unit['opaqElementName'] = sheet['Element name'].tolist()[1:]
    unit['externalWallArea'] = sheet['External wall area'].tolist()[1:]
    unit['externalWallUvalue'] = sheet['External wall U-value'].tolist()[1:]
    unit['shelteredWallArea'] = sheet['Sheltered wall area'].tolist()[1:]
    unit['shelteredWallUvalue'] = sheet['Sheltered wall U-value'].tolist()[1:]
    unit['shelterFactor'] = sheet['Sheltered wall shelter factor'].tolist()[1:]
    unit['partylWallArea'] = sheet['Party wall area'].tolist()[1:]
    unit['externalRoofArea'] = sheet['External roof area'].tolist()[1:]
    unit['externalRoofUvalue'] = sheet['External roof U-value'].tolist()[1:]
    unit['externalRoofType'] = sheet['External roof type'].tolist()[1:]
    unit['externalRoofShelterFactor'] = sheet['External roof shelter factor'].tolist()[1:]
    unit['heatLossFloorArea'] = sheet['Heat loss floor area'].tolist()[1:]
    unit['heatLossFloorUvalue'] = sheet['Heat loss floor U-value'].tolist()[1:]
    unit['heatLossFloorType'] = sheet['Heat loss floor type'].tolist()[1:]
    unit['heatLossFloorShelterFactor'] = sheet['Heat loss floor shelter factor'].tolist()[1:]
    unit['partyCeilingArea'] = sheet['Party ceiling area'].tolist()[1:]
    unit['partyFloorArea'] = sheet['Party floor area'].tolist()[1:]

    # Check if all element types have a name assigned
    filtered_op_elem_type = sheet['Element type'][sheet['Element type'].notna()]
    filtered_op_elem_name = sheet['Element name'][sheet['Element name'].notna()]
    if len(filtered_op_elem_type)!=len(filtered_op_elem_name):
        raise ErrorFound(f'''Error for {unit["propertyName"]}:
                         There is a mismatch between the number of
                         "Element type" entries and the assigned "Element name"''')

    # Check if opaque elements have unique names assigned
    check_op_element_names(unit, sheet)

    # Checking if number of area inputs matches the number of entries of element type
    op_elements = [
        'External wall',
        'Sheltered wall',
        'Party wall',
        'External roof',
        'Heat loss floor',
        'Party ceiling',
        'Party floor'
    ]
    op_elements_titles = [
        'External wall length',
        'Sheltered wall length',
        'Party wall length',
        'External roof area',
        'Heat loss floor area',
        'Party ceiling area',
        'Party floor area'
    ]
    for this_id, element in enumerate(op_elements):
        filtered_elements = sheet['Element type'][sheet['Element type'] == element]
        if len(filtered_elements) != len(sheet[op_elements_titles[this_id]][sheet[op_elements_titles[this_id]].notna()]):
            raise ErrorFound(f'''Error for {unit["propertyName"]}:
                             There is a mismatch between the number of
                             "{element}" entries and their corresponding properties''')

    # Opening types
    unit['openTypeName'] = sheet['Opening type name'].tolist()[1:]
    # unit['openTypeName'] = sheet['Opening type name'][sheet['Opening type name'].notna()].tolist() # THIS IS NOT APPLICABLE AT THIS STAGE BECAUSE THE MATCH_XML FUNCTION REQUIRES A DATAFRAME WITH ALL OPENING PROPERTIES WITH SAME LENGTH
    unit['openingType'] = sheet['Type'].tolist()[1:]
    unit['glzgType'] = sheet['Glazing type'].tolist()[1:]
    unit['uVal'] = sheet['U-value'].tolist()[1:]
    unit['gVal'] = sheet['Solar transmittance'].tolist()[1:]
    unit['frameFactor'] = sheet['Frame factor'].tolist()[1:]

    inputs_window = ['Type','U-value','Solar transmittance','Frame factor']
    inputs_door = ['Type','U-value','Frame factor']
    main = 'Opening type name'
    check_opening_type_data(unit, sheet, main, inputs_window, this_type='Window')
    check_opening_type_data(unit, sheet, main, inputs_door, this_type='Door')

    # Openings
    unit['openLevel'] = sheet['Opening level ref.'].tolist()[1:]
    unit['openName'] = sheet['Opening name'].tolist()[1:]
    unit['openType'] = sheet['Opening type'].tolist()[1:]
    unit['parentElem'] = sheet['Belongs to opaque element'].tolist()[1:]
    unit['openOrientation'] = sheet['Orientation'].tolist()[1:]
    unit['openArea'] = sheet['Area'].tolist()[1:]

    inputs_openings= [
        'Opening level ref.',
        'Opening type',
        'Belongs to opaque element',
        'Orientation',
        'Width',
        'Height',
        'Area',
        'Floor to ceiling?'
    ]
    main = 'Opening name'
    check_openings_data(unit, sheet, main, inputs_openings)

    # Thermal bridges
    thermal_bridges = unit['thermalBridges'] = {}
    for tb, _ in TBs.items():
        thermal_bridge = {}
        if sheet[tb][0] == 'ERROR':
            raise ErrorFound(f'Error for {unit["propertyName"]}: Psi value not entered for thermal bridge {tb}')
        thermal_bridge['psi'] = sheet[tb][sheet[tb].notna()].tolist()[0]
        thermal_bridge['lengths']  = sheet[tb][2:][sheet[tb].notna()].tolist()
        thermal_bridges[tb] = thermal_bridge

    # Mechanical ventilation
    mech_vent_list = [
        'Mech vent present',
        'Ventilation data type',
        'Mech vent type',
        'Vent brand model',
        'MVHR SFP','MVHR HR',
        'Wet rooms',
        'System location',
        'Duct insulation',
        'Duct installation specs',
        'Duct type',
        'Air permeability @50Pa'
    ]
    check_missing_data(unit, sheet, mech_vent_list)

    # Lighting
    lighting_list = [
        'Lighting name',
        'Efficacy',
        'Power',
        'Capacity',
        'Count'
    ]
    check_missing_data(unit, sheet, lighting_list)

    # Heat networks
    heat_networks_list = [
        'Heating network type',
        'Distribution loss space',
        'Heating source 1 - source',
        'Fuel type',
        'Distribution loss',
        'Heating controls',
        'Percentage of heat',
        'Overall efficiency',
        'Heating use'
    ]
    check_missing_data(unit, sheet, heat_networks_list)

    # Water heating
    water_heating_list = [
        'Water heating',
        'Cold water source',
        'Bath count',
        'Shower type',
        'Shower flowrate',
        'Storage type'
    ]
    check_missing_data(unit, sheet, water_heating_list)

    # PV
    PV_list = [
        'PV present?',
        'PV type',
        'Cells peak',
        'PV orientation',
        'PV elevation',
        'PV overshading'
    ]
    check_missing_data(unit, sheet, PV_list)

    return unit

def match_xml(input_unit):
    """Begins matchings the dict format to a nested dictionaries for easier export to XML"""
    # Instantiate output_data dict
    output_data = {}

    # Output data for general unit info
    assessment = output_data['Assessment'] = {}
    assessment['Reference'] = input_unit['propertyName']
    assessment['DwellingOrientation'] = input_unit['Dwelling orientation']
    assessment['CalculationType'] = input_unit['Calculation type']
    assessment['Tenure'] = 'ND'
    assessment['TransactionType'] = int(6)
    assessment['TerrainType'] = input_unit['Terrain type']
    assessment['SimpleComplianceScotland'] = 'false'
    assessment['PropertyType1'] = input_unit['Property type 1']
    assessment['PropertyType2'] = input_unit['Property type 2']
    assessment['PositionOfFlat'] = input_unit['Position of flat']
    assessment['FlatWhichFloor'] = int(input_unit['Which floor'])
    assessment['StoreysInBlock'] = int(input_unit['Tot no. storeys in block'])
    assessment['Storeys'] = len([1 for x in input_unit['floorToSlab'] if x > 0])
    assessment['DateBuilt'] = int(input_unit['Date built'])
    assessment['PropertyAgeBand'] = 'replace_xsi:nul'
    assessment['ShelteredSides'] = int(input_unit['Sheltered sides'])
    assessment['SunlightShade'] = input_unit['Sunlight/sunshade']
    assessment['Basement'] = 'false'
    assessment['LivingArea'] = input_unit['Living area']
    assessment['ThermalMass'] = 'EnterTmpValue'
    assessment['ThermalMassValue'] = input_unit['Thermal mass parameter']
    assessment['LowestFloorHasUnheatedSpace'] = 'replace_xsi:nul'
    assessment['UnheatedFloorArea'] = 'replace_xsi:nul'

    # Output data for measurements
    measurements = assessment['Measurements'] = {}

    # Looping through 10 measurements (storeys), as expected by the XML input in Elmhurst
    # If the storey is not present in the sheet, the output will be just all Os.
    filtered_heat_loss_perim = [x for x in input_unit['heatLossPerim'] if not math.isnan(x)]
    filtered_heated_int_area = [x for x in input_unit['heatedIntArea'] if not math.isnan(x)]
    filtered_floor_to_slab = [x for x in input_unit['floorToSlab'] if not math.isnan(x)]

    if len(filtered_heat_loss_perim) == len(filtered_heated_int_area) == len(filtered_floor_to_slab):
        pass
    else:
        raise ErrorFound(f'''Error for {input_unit["propertyName"]}: 
                        One or multiple inputs among ["Floor to slab", "Heat loss perimeter", 
                        "Heated internal area"] have not been entered''')

    for msrmt in range(9):
        measurement = {}
        if msrmt>0:
            try:
                if math.isnan(input_unit['heatLossPerim'][msrmt-1]):
                    measurement['Storey'] = msrmt
                    measurement['InternalPerimeter'] = 0
                    measurement['InternalFloorArea'] = 0
                    measurement['StoreyHeight'] = 0
                else:
                    measurement['Storey'] = msrmt
                    measurement['InternalPerimeter'] = input_unit['heatLossPerim'][msrmt-1]
                    measurement['InternalFloorArea'] = input_unit['heatedIntArea'][msrmt-1]
                    measurement['StoreyHeight'] = input_unit['floorToSlab'][msrmt-1]
            except:
                measurement['Storey'] = msrmt
                measurement['InternalPerimeter'] = 0
                measurement['InternalFloorArea'] = 0
                measurement['StoreyHeight'] = 0
        else: #for some reason, Elmhurst expects an empty Storey 0
            measurement['Storey'] = msrmt
            measurement['InternalPerimeter'] = 0
            measurement['InternalFloorArea'] = 0
            measurement['StoreyHeight'] = 0
        measurements[f'Measurement{msrmt}']  = measurement

    # Instantiate empty dictionaries for opqaue elements
    ext_walls = assessment['ExternalWalls'] = {}
    party_walls = assessment['PartyWalls'] = {}
    assessment['InternalPartitions'] = []
    ext_roofs = assessment['ExternalRoofs'] = {}
    party_roofs = assessment['PartyRoofs'] = {}
    assessment['InternalCeilings'] = []
    heatloss_floors = assessment['HeatlossFloors'] = {}
    party_floors = assessment['PartyFloors'] = {}
    assessment['InternalFloors'] = []

    op_elements = [
        'opaqElementType',
        "externalWallArea",
        "externalWallUvalue",
        "shelteredWallArea",
        "shelteredWallUvalue",
        "shelterFactor",
        "partylWallArea",
        "externalRoofArea",
        "externalRoofUvalue",
        "externalRoofType",
        "externalRoofShelterFactor",
        "heatLossFloorArea",
        "heatLossFloorUvalue",
        "heatLossFloorType",
        "heatLossFloorShelterFactor",
        "partyCeilingArea",
        "partyFloorArea"
    ]
    op_elements_dict = {col: input_unit[col] for col in op_elements}
    op_elements_df = pd.DataFrame.from_dict(op_elements_dict)

    # Looping through the opaque elements and checking what element type it is.
    # Each type has different outputs
    for this_id, this_type in enumerate(input_unit['opaqElementType']):
        row = op_elements_df.loc[this_id]
        if this_type == 'External wall':
            filtered_row = row[:2][row[:2].notna()]
            if len(filtered_row)==2:
                ext_wall = {}
                ext_wall['Description'] = input_unit['opaqElementName'][this_id]
                ext_wall['Construction'] = 'Other'
                ext_wall['Kappa'] = 0
                ext_wall['GrossArea'] = round(input_unit['externalWallArea'][this_id],3)
                ext_wall['Uvalue'] = input_unit['externalWallUvalue'][this_id]
                ext_wall['ShelterFactor'] = 0
                ext_wall['ShelterCode'] = None
                ext_wall['Type'] = 'Cavity'
                ext_wall['AreaCalculationType'] = 'Gross'
                ext_wall['OpeningsArea'] = 'replace_xsi:nul'
                ext_wall['NettArea'] = 0
                ext_walls[f'ExternalWall{this_id}'] = ext_wall
            else:
                raise ErrorFound(f'Error for {input_unit["propertyName"]}: An "External wall" has been entered without corresponding properties')
        elif this_type == 'Sheltered wall':
            filtered_row = row[2:6][row[2:6].notna()]
            if len(filtered_row)==3:
                shelt_wall = {}
                shelt_wall['Description'] = input_unit['opaqElementName'][this_id]
                shelt_wall['Construction'] = 'Other'
                shelt_wall['Kappa'] = 0
                shelt_wall['GrossArea'] = round(input_unit['shelteredWallArea'][this_id],3)
                shelt_wall['Uvalue'] = input_unit['shelteredWallUvalue'][this_id]
                shelt_wall['ShelterFactor'] = input_unit['shelterFactor'][this_id]
                shelt_wall['ShelterCode'] = None
                shelt_wall['Type'] = 'Cavity'
                shelt_wall['AreaCalculationType'] = 'Gross'
                shelt_wall['OpeningsArea'] = 'replace_xsi:nul'
                shelt_wall['NettArea'] = 0
                ext_walls[f'ExternalWall{this_id}'] = shelt_wall
            else:
                raise ErrorFound(f'Error for {input_unit["propertyName"]}: A "Sheltered wall" has been entered without corresponding properties')
        elif this_type == 'Party wall':
            if row.iloc[6]>0:
                party_wall = {}
                party_wall['Description'] = input_unit['opaqElementName'][this_id]
                party_wall['Construction'] = 'Other'
                party_wall['Kappa'] = 0
                party_wall['GrossArea'] = round(input_unit['partylWallArea'][this_id],3)
                party_wall['Uvalue'] = 0
                party_wall['ShelterFactor'] = 0
                party_wall['ShelterCode'] = None
                party_wall['Type'] = 'FilledWithEdge'
                party_walls[f'PartyWall{this_id}'] = party_wall
            else:
                raise ErrorFound(f'Error for {input_unit["propertyName"]}: A "Party wall" has been entered without corresponding properties')


        elif this_type == 'External roof':
            filtered_row = row[7:11][row[7:11].notna()]
            if len(filtered_row)==4:
                ext_roof = {}
                ext_roof['Description'] = input_unit['opaqElementName'][this_id]
                try:
                    ext_roof['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][this_id]-1))]
                except Exception as exc:
                    raise ErrorFound(f'The level reference entered for {this_type} is not listed under "Levels"') from exc
                ext_roof['Construction'] = 'Other'
                ext_roof['Kappa'] = 0
                ext_roof['GrossArea'] = input_unit['externalRoofArea'][this_id]
                ext_roof['Type'] = input_unit['externalRoofType'][this_id]
                ext_roof['UValue'] = input_unit['externalRoofUvalue'][this_id]
                ext_roof['ShelterFactor'] = input_unit['externalRoofShelterFactor'][this_id]
                ext_roof['ShelterCode'] = None
                ext_roof['AreaCalculationType'] = 'Gross'
                ext_roof['OpeningsArea'] = 'replace_xsi:nul'
                ext_roof['NettArea'] = 0
                ext_roofs[f'ExternalRoof{this_id}'] = ext_roof
            else:
                raise ErrorFound(f'Error for {input_unit["propertyName"]}: An "External roof" has been entered without corresponding properties')
        elif this_type == 'Heat loss floor':
            filtered_row = row[11:15][row[11:15].notna()]
            if len(filtered_row)==4:
                heatloss_floor = {}
                heatloss_floor['Description'] = input_unit['opaqElementName'][this_id]
                heatloss_floor['Construction'] = 'Other'
                heatloss_floor['Kappa'] = 0
                heatloss_floor['Area'] = input_unit['heatLossFloorArea'][this_id]
                try:
                    heatloss_floor['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][this_id]-1))]
                except Exception as exc:
                    raise ErrorFound(f'The level reference entered for {this_type} is not listed under "Levels"') from exc
                heatloss_floor['Type'] = input_unit['heatLossFloorType'][this_id]
                heatloss_floor['UValue'] = input_unit['heatLossFloorUvalue'][this_id]
                heatloss_floor['ShelterFactor'] = input_unit['heatLossFloorShelterFactor'][this_id]
                heatloss_floor['ShelterCode'] = None
                heatloss_floors[f'HeatLossFloor{this_id}'] = heatloss_floor
            else:
                raise ErrorFound(f'Error for {input_unit["propertyName"]}: A "Heat loss floor" has been entered without corresponding properties')
        elif this_type == 'Party ceiling':
            if row.iloc[15]>0:
                party_roof = {}
                party_roof['Description'] = input_unit['opaqElementName'][this_id]
                try:
                    party_roof['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][this_id]-1))]
                except Exception as exc:
                    raise ErrorFound(f'The level reference entered for {this_type} is not listed under "Levels"') from exc
                party_roof['Construction'] = 'Other'
                party_roof['Kappa'] = 0
                party_roof['GrossArea'] = input_unit['partyCeilingArea'][this_id]
                party_roofs[f'Roof{this_id}'] = party_roof
            else:
                raise ErrorFound(f'Error for {input_unit["propertyName"]}: A "Party ceiling" has been entered without corresponding properties')
        elif this_type == 'Party floor':
            if row.iloc[16]>0:
                party_floor = {}
                party_floor['Description'] = input_unit['opaqElementName'][this_id]
                party_floor['Construction'] = 'Other'
                party_floor['Kappa'] = 0
                party_floor['Area'] = input_unit['partyFloorArea'][this_id]
                try:
                    party_floor['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][this_id]-1))]
                except Exception as exc:
                    raise ErrorFound(f'The level reference entered for {this_type} is not listed under "Levels"') from exc
                party_floors[f'Floor{this_id}'] = party_floor
            else:
                raise ErrorFound(f'Error for {input_unit["propertyName"]}: A "Party floor" has been entered without corresponding properties')
    # Misc objects that need to be included in the XML for Elmhurst
    # (but currently are not allowed to be entered in the excel sheet)
    assessment['ThermalBridgesCalculation'] = 'CalculateBridges'
    assessment['ThermalBridgingSpreadsheet'] = 'Summary'
    assessment['ThermalBridgesYvalue'] = 0
    assessment['ThermalBridgesDescription'] = []
    assessment['PointThermalBridgingX'] = 'replace_xsi:nul'
    assessment['OpenChimneys'] = 0
    assessment['OpenFlues'] = 0
    assessment['ChimneysFluesClosedFire'] = 0
    assessment['FluesSolidFuelBoiler'] = 0
    assessment['FluesOtherHeater'] = 0
    assessment['BlockedChimneys'] = 0
    assessment['IntermittentFans'] = 0
    assessment['PassiveVents'] = 0
    assessment['FluelessGasFires'] = 0
    assessment['NoFixedLighting'] = 'false'
    assessment['LightingCapacityCalculation'] = 'replace_xsi:nul'

    # Output data for lighting
    lightings = assessment['Lightings'] = {}
    lighting = lightings['Lighting'] = {}
    lighting['Name'] = input_unit['Lighting name']
    lighting['Efficacy'] = input_unit['Efficacy']
    lighting['Power'] = int(input_unit['Power'])
    lighting['Capacity'] = int(input_unit['Capacity'])
    lighting['Count'] = int(input_unit['Count'])

    # Misc objects that need to be included in the XML for Elmhurst
    # (but currently are not allowed to be entered in the excel sheet)
    assessment['ElectricityTariff'] ='Standard'
    assessment['SmartElectricityMeterFitted'] = 'false'
    assessment['SmartGasMeterFitted'] = 'false'
    assessment['SolarPanelPresent'] = 'false'
    assessment['PressureTest'] = 'true'
    assessment['PressureTestMethod'] = 'BlowerDoor'
    assessment['Designed_AP50_AP4'] = input_unit['Air permeability @50Pa']
    assessment['AsBuilt_AP50_AP4'] = 0
    assessment['PropertyTested'] = 'true'
    assessment['SmokeControlArea'] = 'Unknown'
    assessment['ThermallySeparated'] = 'NoConservatory'
    assessment['DraughtProofing'] = 100
    assessment['DraughtLobby'] = 'false'
    assessment['Floor1AreaCalculated'] = 'false'
    assessment['PhotovoltaicUnitApportionedEnergy'] = 'replace_xsi:nul'
    assessment['ConnectedToDwelling'] = 'Yes'
    assessment['Diverter'] = 'No'
    assessment['BatteryCapacity'] = 0

    # Output data for PV panels. Checks if PVs are present, otherwise retuns empty tag
    if input_unit['PV present?'] == 'Yes':
        assessment['PhotovoltaicUnitType'] = input_unit['PV type']
        pvs = assessment['PhotovoltaicUnits'] = {}
        pv = pvs['PhotovoltaicUnit'] = {}
        pv['CellsPeak'] = input_unit['Cells peak']
        pv['Orientation'] = input_unit['PV orientation']
        pv['Elevation'] = input_unit['PV elevation']
        pv['Overshading'] = input_unit['PV overshading']
        pv['Fghrs'] = 'false'
        pv['MCSCertificate'] = 'false'
        pv['OvershadingFactor'] = 0
    else:
        assessment['PhotovoltaicUnitType'] = None
        assessment['PhotovoltaicUnits'] = {}

    # Output data for opening types
    assessment['OpeningTypes'] = {}
    opening_types = assessment['OpeningTypes'] = {}

    # Looping through each opening type and only including if data is entered
    for this_id, name in enumerate(input_unit['openTypeName']):
        if input_unit['uVal'][this_id]>0:
            opening_type = {}
            opening_type['Description'] = name
            opening_type['DataSource'] = 'Manufacturer'
            opening_type['Type'] = input_unit['openingType'][this_id]
            if input_unit['openingType'][this_id]=='Window':
                opening_type['Glazing'] = 'Double'
                opening_type['GlazingGap'] = 'replace_xsi:nul'
                opening_type['GlazingFillingType'] = None
                opening_type['SolarTrans'] = input_unit['gVal'][this_id]
            else:
                opening_type['Glazing'] = 'replace_xsi:nul'
                opening_type['GlazingGap'] = 'replace_xsi:nul'
                opening_type['GlazingFillingType'] = None
                opening_type['SolarTrans'] = 0
            opening_type['FrameType'] = 'Wood'
            opening_type['FrameFactor'] = input_unit['frameFactor'][this_id]
            opening_type['UValue'] = input_unit['uVal'][this_id]
            opening_types[f'OpeningType{this_id}'] = opening_type

    # Output data for openings
    openings = assessment['Openings'] = {}

    # Looping through each opening and only including if data is entered
    for this_id, name in enumerate(input_unit['openName']):
        # Write inputs only if area is entered
        if input_unit['openArea'][this_id]>0:
            # Check for refernece levels that have not been listed in "levels"
            try:
                levels_naming[str(int(input_unit['openLevel'][this_id]-1))]
            except Exception as exc:
                raise ErrorFound(f'The level reference entered for opening "{name}" is not listed under "Levels"') from exc

            opening = {}
            opening['this_id'] = this_id
            for this_ido,this_type in enumerate(input_unit['openTypeName']):
                if input_unit['openType'][this_id] == this_type:
                    opening['OpeningTypeIndex'] = this_ido
            opening['Description'] = name
            opening['LocationBuildingPartIndex'] = 0

            counter = 0
            wall_list = []
            for this_ido, this_type in enumerate(input_unit['opaqElementType']):
                if this_type == "External wall" or this_type == "Sheltered wall":
                    wall_list.append(input_unit['opaqElementName'][this_ido])
            for this_ido,parent in enumerate(wall_list):
                if input_unit['parentElem'][this_id] == parent:
                    opening['LocationWallIndex'] = this_ido
                    counter+=1
            if counter == 0:
                raise ErrorFound(f'''The parent element "{input_unit["parentElem"][this_id]}" 
                                referred by the "{name}" opening element does not exist 
                                or it is not an External or Sheltered wall''')
            opening['LocationRoofIndex'] = 'replace_xsi:nul'
            opening['Orientation'] = input_unit['openOrientation'][this_id]
            opening['AreaType'] = 'Total'
            opening['AreaScaleType'] = 'Meters'
            opening['Area'] = input_unit['openArea'][this_id]
            opening['AreaRecCalculation'] = []
            opening['RoofLightsPitch'] = 0
            openings[f'Opening{this_id}'] = opening

    # Output data for thermal bridges
    thermal_bridges = assessment['ThermalBridges'] = {}

    # Looping through each TB and only including if data is entered
    for tb,name in TBs.items():
        lengths = input_unit['thermalBridges'][tb]['lengths']
        psi = input_unit['thermalBridges'][tb]['psi']
        for this_id,length in enumerate(lengths):
            thermal_bridge = {}
            thermal_bridge['TypeSource'] = 'IndependentlyAssessed'
            thermal_bridge['Length'] = length
            thermal_bridge['PsiValue'] = psi
            thermal_bridge['K1Index'] = name
            thermal_bridge['Imported'] = 'False'
            thermal_bridge['Adjusted'] = psi
            thermal_bridge['Reference'] = []
            thermal_bridges[f'ThermalBridge-{tb}-{this_id}'] = thermal_bridge

    # output data for mech vent
    if input_unit['Mech vent present'] == "Yes":
        mechvent = assessment['MechanicalVentilation'] = {}
        mechvent['DataType'] = input_unit['Ventilation data type']
        mechvent['Type'] = input_unit['Mech vent type']
        mechvent['PcdfIndex'] = 'replace_xsi:nul'
        mechvent['PcdfItem'] = 'replace_xsi:nul'
        mechvent['ManufacturerSFP'] = input_unit['MVHR SFP']
        mechvent['DuctType'] = input_unit['Duct type']
        mechvent['WetRooms'] = int(input_unit['Wet rooms'])
        mechvent['BrandModel'] = input_unit['Vent brand model']
        mechvent['MVHRDuctInsulated'] = 'replace_xsi:nul'
        mechvent['DuctInsulation'] = 'replace_xsi:nul'
        mechvent['MVHREfficiency'] = input_unit['MVHR HR']
        mechvent['ApprovedInstallation'] = "false"
        mechvent['SFPFromInstallerCertificate'] = "false"
        mechvent['MVHRSystemLocation'] = input_unit['System location']
        if mechvent['MVHRSystemLocation'] == "Outside":
            mechvent['MVHRDuctInsulated'] = input_unit['Duct insulation']
        mechvent['DuctInsulationLevel'] = input_unit['Duct installation specs']

    assessment['MechanicalVentilationDecentralised'] = []
    assessment['HeatingsInteraction'] = 'SeparatePartsOfHouse'

    # Output data for the main heating systems
    # These are empty nested items that Elmhurst requires as input
    for mhs in range(2):
        main_heating_system = assessment[f'MainHeatingSystem{mhs+1}'] = {}
        main_heating_system['HeatingDataType'] = 'None'
        main_heating_system['Fraction'] = 0
        main_heating_system['PcdfIndex'] = 0
        main_heating_system['BoilerEfficiencyType'] = 'replace_xsi:nul'
        main_heating_system['EfficiencyWinter'] = 0
        main_heating_system['EfficiencySummer'] = 0
        main_heating_system['TestMethod'] = 'replace_xsi:nul'
        main_heating_system['MHSCtrlPcdfIndex'] = 'replace_xsi:nul'
        main_heating_system['CompensatorPcdfIndex'] = 'replace_xsi:nul'
        main_heating_system['HetasApprovedSystem'] = 'false'
        main_heating_system['FlueType'] = 'replace_xsi:nul'
        main_heating_system['FanAssistedFlue'] = 'false'
        main_heating_system['McsCertificate'] = 'false'
        main_heating_system['Pumped'] = 'replace_xsi:nul'
        main_heating_system['HeatingPumpAge'] = 'replace_xsi:nul'
        main_heating_system['OilPumpInside'] = 'false'
        main_heating_system['HeatEmitter'] = 'replace_xsi:nul'
        main_heating_system['UnderfloorHeating'] = 'replace_xsi:nul'
        main_heating_system['CombiType'] = 'replace_xsi:nul'
        main_heating_system['CombiKeepHotType'] = 'replace_xsi:nul'
        main_heating_system['CombiStoreType'] = 'replace_xsi:nul'
        main_heating_system['ElectricCPSUtemperature'] = 'replace_xsi:nul'
        main_heating_system['FIcase'] = 'replace_xsi:nul'
        main_heating_system['FIwater'] = 'replace_xsi:nul'
        main_heating_system['BurnerControl'] = 'replace_xsi:nul'
        main_heating_system['DelayedStartStat'] = 'false'
        main_heating_system['FlowTemperature'] = 'replace_xsi:nul'
        main_heating_system['BoilerInterlock'] = 'false'
        main_heating_system['StorageHeaters'] = {}
        main_heating_system['FlowTemperatureValue'] = 'replace_xsi:nul'
        main_heating_system['SapCode'] = 'replace_xsi:nul'
        main_heating_system['FuelType'] = 'replace_xsi:nul'
        main_heating_system['CtrlSapCode'] = 'replace_xsi:nul'

    # Output data for the secondary heating systems.
    # This is an empty item that Elmhurst requires as input
    secondary_heating = assessment['SecondaryHeating'] = {}
    secondary_heating['HeatingDataType'] = 'None'
    secondary_heating['TestMethod'] = 'replace_xsi:nul'
    secondary_heating['HetasApprovedSystems'] = 'false'
    secondary_heating['Efficiency'] = 'replace_xsi:nul'
    secondary_heating['SapCode'] = 0
    secondary_heating['FuelType'] = 'replace_xsi:nul'

    # Output data for community heating
    community_heating = assessment['CommunityHeating'] = {}
    community_heating['Type'] = input_unit['Heating network type']
    community_heating['DistributionLossSpace'] = input_unit['Distribution loss space']
    community_heating['DistributionLossWater'] = 'replace_xsi:nul'
    community_heating['ChargingLinked'] = 'replace_xsi:nul'
    heat_source = community_heating['HeatSource'] = {}
    # Elmhurst expects 5 heat sources as input
    for chs in range(5):
        comm_heat_source = heat_source[f'CommunityHeatSource{chs+1}'] = {}
        # Only inputting into the first heat source, the rest is kept empty
        if chs == 0:
            comm_heat_source['Source'] = input_unit['Heating source 1 - source']
            comm_heat_source['Fraction'] = input_unit['Percentage of heat']
            comm_heat_source['FuelType'] = input_unit['Fuel type']
            comm_heat_source['OveralEfficiency'] = input_unit['Overall efficiency']
            comm_heat_source['HeatPowerRatio'] = 'replace_xsi:nul'
            comm_heat_source['ElectricalEfficiency'] = 'replace_xsi:nul'
            comm_heat_source['HeatEfficiency'] = 'replace_xsi:nul'
            comm_heat_source['HeatingUse'] = input_unit['Heating use']
            comm_heat_source['CHPFuelFactor'] = 'replace_xsi:nul'
            comm_heat_source['EfficiencyType'] = 'replace_xsi:nul'
        else:
            comm_heat_source['Source'] = 'None'
            comm_heat_source['Fraction'] = 'replace_xsi:nul'
            comm_heat_source['FuelType'] = 'replace_xsi:nul'
            comm_heat_source['OveralEfficiency'] = 'replace_xsi:nul'
            comm_heat_source['HeatPowerRatio'] = 'replace_xsi:nul'
            comm_heat_source['ElectricalEfficiency'] = 'replace_xsi:nul'
            comm_heat_source['HeatEfficiency'] = 'replace_xsi:nul'
            comm_heat_source['HeatingUse'] = input_unit['Heating use']
            comm_heat_source['CHPFuelFactor'] = 'replace_xsi:nul'
            comm_heat_source['EfficiencyType'] = 'replace_xsi:nul'

    community_heating['DistributionLossSpaceValue'] = input_unit['Distribution loss']
    community_heating['DistributionLossWaterValue'] = 'replace_xsi:nul'
    community_heating['SpacePCDFIndex'] = 'replace_xsi:nul'
    community_heating['WaterPCDFIndex'] = 'replace_xsi:nul'
    community_heating['CtrlSapCode'] = int(input_unit['Heating controls'])
    community_heating['ExistingSpace'] = 'replace_xsi:nul'
    community_heating['ExistingWater'] = 'replace_xsi:nul'
    community_heating['UseNotionalSpace'] = 'replace_xsi:nul'
    community_heating['UseNotionalWater'] = 'replace_xsi:nul'

    # Output data for water heating systems DHW
    water_heating_system = assessment['WaterHeatingSystem'] = {}
    water_heating_system['WaterHeatingType'] = input_unit['Water heating']
    water_heating_system['LowWaterUse'] = 'false'
    water_heating_system['ImmersionHeaterType'] = 'replace_xsi:nul'
    water_heating_system['SummerImmersion'] = 'false'
    water_heating_system['SuplementaryImmersion'] = 'false'
    water_heating_system['ImmersionOnlyHeatingHotWater'] = 'false'
    water_heating_system['ThermalStore'] = 'None'
    water_heating_system['ThermalStorePipework'] = 'replace_xsi:nul'
    water_heating_system['HotWaterCylinder'] = input_unit['Storage type']
    water_heating_system['InsulationType'] = 'replace_xsi:nul'
    water_heating_system['InsulationThickness'] = 'replace_xsi:nul'
    water_heating_system['InsulationThicknessType'] = 'replace_xsi:nul'
    water_heating_system['Volume'] = 'replace_xsi:nul'
    water_heating_system['CylinderStat'] = 'false'
    water_heating_system['PipeworkInsulation'] = 'replace_xsi:nul'
    water_heating_system['InHeatedSpace'] = 'false'
    water_heating_system['InAiringCupboard'] = 'false'
    water_heating_system['SeparateTimeControl'] = 0
    water_heating_system['LossFactor'] = 1.46
    water_heating_system['SolarPanelType'] = 'replace_xsi:nul'
    water_heating_system['SolarAreaType'] = 'Aperture'
    water_heating_system['SolarArea'] = 0
    water_heating_system['SolarNi'] = 0
    water_heating_system['SolarA1'] = 0
    water_heating_system['SolarA2'] = 0
    water_heating_system['SolarAGRatio'] = 0
    water_heating_system['SolarLoopEfficiency'] = 0.9
    water_heating_system['SolarKhem'] = 0
    water_heating_system['SolarHeatLossCoeff'] = 'replace_xsi:nul'
    water_heating_system['SolarIsFromCommunity'] = 'false'
    water_heating_system['SolarServiceProvision'] = 'replace_xsi:nul'
    water_heating_system['SolarPanelOrientation'] = 'replace_xsi:nul'
    water_heating_system['SolarElevation'] = 'replace_xsi:nul'
    water_heating_system['SolarOvershadingType'] = 'replace_xsi:nul'
    water_heating_system['SolarVolume'] = 'replace_xsi:nul'
    water_heating_system['SolarPumpElectricallyPowered'] = 'false'
    water_heating_system['SolarCombinedCylinder'] = 'false'
    water_heating_system['ColdWaterSource'] = input_unit['Cold water source']
    water_heating_system['BathCount'] = int(input_unit['Bath count'])
    water_heating_system['WWHRSBathCount'] = 'replace_xsi:nul'
    water_heating_system['SapCode'] = 901
    water_heating_system['FuelType'] = 'replace_xsi:nul'
    water_heating_system['HIUPcdfIndex'] = 'replace_xsi:nul'

    # Output data for showers (only one shower item is exported)
    showers = assessment['Showers'] = {}
    shower = showers['Shower'] = {}
    shower['Description'] = 'Shower1'
    shower['ShowerType'] = input_unit['Shower type']
    shower['FlowRate'] = input_unit['Shower flowrate']
    shower['RatedPower'] = round(1.1625 * input_unit['Shower flowrate'],1)
    shower['Connected'] = 'true'
    shower['ConnectedTo'] = 'Storage'

    # Misc objects that need to be included in the XML for Elmhurst
    # (but currently are not allowed to be entered in the excel sheet)
    assessment['WindTurbineType'] = 'None'
    assessment['WindTurbines'] = []
    assessment['ApportionedEnergy'] = 'replace_xsi:nul'
    assessment['SpecialTechnology'] = []
    assessment['RelatedPartyDisclosure'] = 1
    assessment['Recomm_N'] = 'true'
    assessment['Recomm_U'] = 'true'
    assessment['Recomm_V2'] = 'true'
    assessment['IATSReference'] = []
    assessment['IATSDataExists'] = 'false'
    assessment['IATSTestDate'] = 'replace_xsi:nul'
    assessment['ExportCapableMeter'] = 'false'
    assessment['ShowSpaceHeatDemand'] = 'replace_xsi:nul'
    plot = output_data['Plot'] = {}
    plot['Reference'] = input_unit['propertyName']
    plot['TypeReference'] = input_unit['propertyName']
    plot['RegsRegion'] = 'England'
    plot['Region'] = 'Thames'
    plot['HouseName'] = []
    plot['HouseNumber'] = []
    plot['Postcode'] = []
    plot['Street'] = []
    plot['Town'] = []
    plot['County'] = []
    plot['ClientId'] = 'replace_xsi:nul'
    plot['UPRN'] = []
    plot['AddressLine1'] = []
    plot['AddressLine2'] = []
    plot['AddressLine3'] = []
    plot['TownAsDesigned'] = []
    plot['PostcodeAsDesigned'] = []
    plot['AssessorId'] = '47929'
    plot['Id'] = '231765'
    plot['GroupId'] = '32079'
    plot['SubGroupId'] = 'replace_xsi:nul'
    plot['AssessorCode'] = []
    plot['AssessorTitle'] = []
    plot['AssessorName'] = []
    plot['AssessorSurname'] = []

    return output_data

def find_and_replace(file_path, value = 'replace_xsi:nul'):
    """Writes XML Data Out"""
    with open(file_path, 'r', encoding="utf-8") as file:
        xml_data = file.read()
        root = ET.fromstring(xml_data)
        # Add root attributes
        root.set("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        # Traverse XML and check if values match the flag string. 
        # If they do, remove it and add string as tag attribute
        def traverse(element):
            if element.text == value:
                old_tag = element.tag
                element.tag = f'{old_tag} xsi:nil="true"'
                element.text = ''
            # Check for integers in tags and remove them
            text = str(element.tag)
            for test_text in ['Measurement']:
                for i in range(10):
                    if text == test_text+str(i):
                        element.tag = test_text
            for test_text in [
                    'ExternalWall',
                    'PartyWall',
                    'ExternalRoof',
                    'HeatLossFloor',
                    'Floor',
                    'Roof',
                    'OpeningType',
                    'Opening'
                ]:
                for i in range(20):
                    if text == test_text+str(i):
                        element.tag = test_text
            for test_text in ['CommunityHeatSource']:
                for i in range(5):
                    if text == test_text+str(i+1):
                        element.tag = test_text
            # check_tb = ['ThermalBridgesCalculation','ThermalBridgesYvalue']
            if 'ThermalBridge-' in text:
                element.tag = 'ThermalBridge'
            for child in element:
                traverse(child)
        traverse(root)
        updated_xml_data = ET.tostring(root, encoding='utf-8', method='xml').decode()
    return updated_xml_data
