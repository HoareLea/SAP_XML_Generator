import xml.etree.ElementTree as ET
from xml.dom import minidom
from TBs import TBs
from levels_naming import levels_naming
import math
import pandas as pd

def data_to_xml(dictionary, root_name='AssessmentFull'):
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
    reparsed = minidom.parseString(elem)
    prettified = reparsed.toprettyxml(indent="  ")
    final = prettified.replace('<?xml version="1.0" ?>\n', '')
    return final

def check_missing_general_data(unit, sheet, property):
    print(sheet[property].tolist()[1])
    # if sum(sheet[property].isna())==10:
    if math.isnan(sheet[property].tolist()[1]):
        raise Exception(f'Error for {unit["propertyName"]}: {property} has not been entered')
    unit[property] = sheet[property].tolist()[1]

def input_reader(sheet):
    #instantiate unit dict. this will hold all the inputs from the excel sheet
    unit = {}
    
    # General information
    unit['propertyName'] = sheet['Property name'].tolist()[1]
    check_missing_general_data(unit,sheet,'Dwelling orientation')
    check_missing_general_data(unit,sheet,'Calculation type')
    check_missing_general_data(unit,sheet,'Terrain type')
    check_missing_general_data(unit,sheet,'Property type 1')
    check_missing_general_data(unit,sheet,'Property type 2')
    check_missing_general_data(unit,sheet,'Position of flat')
    check_missing_general_data(unit,sheet,'Which floor')
    check_missing_general_data(unit,sheet,'Tot no. storeys in block')
    check_missing_general_data(unit,sheet,'No. storeys')
    check_missing_general_data(unit,sheet,'Date built')
    check_missing_general_data(unit,sheet,'Sheltered sides')
    check_missing_general_data(unit,sheet,'Sunlight/sunshade')
    check_missing_general_data(unit,sheet,'Thermal mass parameter')
    check_missing_general_data(unit,sheet,'Living area')
    
    # Level information
    unit['floorToSlab'] = sheet['Floor to slab'].tolist()[1:]
    unit['heatedIntArea'] = sheet['Heated internal floor area'].tolist()[1:]
    unit['heatLossPerim'] = sheet['Heat loss perimeter'].tolist()[1:]
    
    # Opaque elements
    unit['opaqElementLevel'] = sheet['Level of opaque element'].tolist()[1:]
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

    #Opening types
    unit['openTypeName'] = sheet['Opening type name'].tolist()[1:]
    unit['openingType'] = sheet['Type'].tolist()[1:]
    unit['glzgType'] = sheet['Glazing type'].tolist()[1:]
    unit['uVal'] = sheet['U-value'].tolist()[1:]
    unit['gVal'] = sheet['Solar transmittance'].tolist()[1:]
    unit['frameFactor'] = sheet['Frame factor'].tolist()[1:]

    # Openings
    unit['openName'] = sheet['Opening name'].tolist()[1:]
    unit['openType'] = sheet['Opening type'].tolist()[1:]
    unit['parentElem'] = sheet['Belongs to opaque element'].tolist()[1:]
    unit['openOrientation'] = sheet['Orientation'].tolist()[1:]
    unit['openArea'] = sheet['Area'].tolist()[1:]

    # Thermal bridges
    thermal_bridges = unit['thermalBridges'] = {}
    for tb, _ in TBs.items():
        thermal_bridge = {}
        thermal_bridge['psi'] = sheet[tb].tolist()[0]
        thermal_bridge['lengths']  = sheet[tb].tolist()[2:]
        thermal_bridges[tb] = thermal_bridge
    
    # Ventilation
    unit['mechVentPresent'] = sheet['Mech vent present'].tolist()[1]
    unit['ventDataType'] = sheet['Ventilation data type'].tolist()[1]
    unit['ventType'] = sheet['Mech vent type'].tolist()[1]
    unit['brandModel'] = sheet['Vent brand model'].tolist()[1]
    unit['MVHR SFP'] = sheet['MVHR SFP'].tolist()[1]
    unit['MVHR HR'] = sheet['MVHR HR'].tolist()[1]
    unit['wetRooms'] = int(sheet['Wet rooms'].tolist()[1])
    unit['systemLocation'] = sheet['System location'].tolist()[1]
    unit['ductInsulated'] = sheet['Duct insulation'].tolist()[1]
    unit['ductInstallationSpecs'] = sheet['Duct installation specs'].tolist()[1]
    unit['ductType'] = sheet['Duct type'].tolist()[1]
    unit['airPerm'] = sheet['Air permeability @50Pa'].tolist()[1]
    
    # Lighting
    unit['lightingName'] = sheet['Lighting name'].tolist()[1]
    unit['lightingEfficacy'] = int(sheet['Efficacy'].tolist()[1])
    unit['lightingPower'] = int(sheet['Power'].tolist()[1])
    unit['lightingCapacity'] = int(sheet['Capacity'].tolist()[1])
    unit['lightingCount'] = int(sheet['Count'].tolist()[1])
    
    # Heat networks
    unit['heatingNetworkType'] = sheet['Heating network type'].tolist()[1]
    unit['distributionLossSpace'] = sheet['Distribution loss space'].tolist()[1]
    unit['heatingSource1'] = sheet['Heating source 1 - source'].tolist()[1]
    unit['fuelType'] = sheet['Fuel type'].tolist()[1]
    unit['distributionLoss'] = sheet['Distribution loss'].tolist()[1]
    unit['heatingControls'] = sheet['Heating controls'].tolist()[1]
    unit['percentageHeat'] = sheet['Percentage of heat'].tolist()[1]
    unit['overallEfficiency'] = sheet['Overall efficiency'].tolist()[1]
    unit['heatingUse'] = sheet['Heating use'].tolist()[1]
    
    # Water heating
    unit['waterHeating'] = sheet['Water heating'].tolist()[1]
    unit['coldWaterSource'] = sheet['Cold water source'].tolist()[1]
    unit['bathCount'] = sheet['Bath count'].tolist()[1]
    unit['showerType'] = sheet['Shower type'].tolist()[1]
    unit['showerFlowrate'] = sheet['Shower flowrate'].tolist()[1]
    unit['storageType'] = sheet['Storage type'].tolist()[1]
    
    # PV
    unit['PVPresent'] = sheet['PV present?'].tolist()[1]
    unit['PVType'] = sheet['PV type'].tolist()[1]
    unit['cellsPeak'] = sheet['Cells peak'].tolist()[1]
    unit['PVOrientation'] = sheet['PV orientation'].tolist()[1]
    unit['PVElevation'] = sheet['PV elevation'].tolist()[1]
    unit['PVOvershading'] = sheet['PV overshading'].tolist()[1]
    
    return unit

def match_xml(input_unit):
    
    # instantiate output_data dict. this will hold all the output data, arranged as nested dictionaries. the nested dictionrary structure allows for an easier export to XML.
    output_data = {}
    
    # output data for general unit info
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
    
    # output data for measurements
    measurements = assessment['Measurements'] = {}
    # looping through 10 measurements (storeys), as expected by the XML input in Elmhurst. if the storey is not present in the sheet, the output will be just all Os.
    filteredHeatLossPerim = [x for x in input_unit['heatLossPerim'] if not math.isnan(x)]
    filteredHeatedIntArea = [x for x in input_unit['heatedIntArea'] if not math.isnan(x)]
    filteredFloorToSlab = [x for x in input_unit['floorToSlab'] if not math.isnan(x)]
    
    if len(filteredHeatLossPerim) == len(filteredHeatedIntArea) == len(filteredFloorToSlab):
        pass
    else:
        raise Exception(f'Error for {input_unit["propertyName"]}: One item among [Floor to slab, Heat loss perimeter, Heated internal area] has not been entered')
    
    for msrmt in range(9):
        measurement = {}
        if msrmt>0:
            if input_unit['heatLossPerim'][msrmt-1]>0:
                measurement['Storey'] = msrmt
                measurement['InternalPerimeter'] = input_unit['heatLossPerim'][msrmt-1]
                measurement['InternalFloorArea'] = input_unit['heatedIntArea'][msrmt-1]
                measurement['StoreyHeight'] = input_unit['floorToSlab'][msrmt-1]
            else:
                measurement['Storey'] = msrmt
                measurement['InternalPerimeter'] = 0
                measurement['InternalFloorArea'] = 0
                measurement['StoreyHeight'] = 0
        else: #for some reason, Elmhurst expects an empty Storey 0
            measurement['Storey'] = msrmt
            measurement['InternalPerimeter'] = 0
            measurement['InternalFloorArea'] = 0
            measurement['StoreyHeight'] = 0
        measurements['Measurement{}'.format(msrmt)]  = measurement
    
    # output data for opqaue elements
    ext_walls = assessment['ExternalWalls'] = {}
    party_walls = assessment['PartyWalls'] = {}
    assessment['InternalPartitions'] = []
    ext_roofs = assessment['ExternalRoofs'] = {}
    party_roofs = assessment['PartyRoofs'] = {}
    assessment['InternalCeilings'] = []
    heatloss_floors = assessment['HeatlossFloors'] = {}
    party_floors = assessment['PartyFloors'] = {}
    assessment['InternalFloors'] = []
    
    op_elements = ['opaqElementType', "externalWallArea","externalWallUvalue","shelteredWallArea",
                   "shelteredWallUvalue","shelterFactor","partylWallArea","externalRoofArea","externalRoofUvalue",
                   "externalRoofType","externalRoofShelterFactor","heatLossFloorArea","heatLossFloorUvalue",
                   "heatLossFloorType","heatLossFloorShelterFactor","partyCeilingArea","partyFloorArea"]
    op_elements_dict = {col: input_unit[col] for col in op_elements}
    op_elements_df = pd.DataFrame.from_dict(op_elements_dict)
    
    # looping through the opaque elemnts and checking what element type it is. each type has a different outputs
    for id, type in enumerate(input_unit['opaqElementType']):
        row = op_elements_df.loc[id]
        # print(row)
        if type == 'External wall':
            if row[1]>0 and row[2]>0:
                ext_wall = {}
                ext_wall['Description'] = input_unit['opaqElementName'][id]
                ext_wall['Construction'] = 'Other'
                ext_wall['Kappa'] = 0
                ext_wall['GrossArea'] = round(input_unit['externalWallArea'][id],3)
                ext_wall['Uvalue'] = input_unit['externalWallUvalue'][id]
                ext_wall['ShelterFactor'] = 0
                ext_wall['ShelterCode'] = None
                ext_wall['Type'] = 'Cavity'
                ext_wall['AreaCalculationType'] = 'Gross'
                ext_wall['OpeningsArea'] = 'replace_xsi:nul'
                ext_wall['NettArea'] = 0
                ext_walls['ExternalWall{}'.format(id)] = ext_wall
            else:
                raise Exception(f'Warning for {input_unit["propertyName"]}: An external wall has been entered without corresponding properties')
        elif type == 'Sheltered wall':
            if row[3]>0 and row[4]>0 and row[5]>=0:
                shelt_wall = {}
                shelt_wall['Description'] = input_unit['opaqElementName'][id]
                shelt_wall['Construction'] = 'Other'
                shelt_wall['Kappa'] = 0
                shelt_wall['GrossArea'] = round(input_unit['shelteredWallArea'][id],3)
                shelt_wall['Uvalue'] = input_unit['shelteredWallUvalue'][id]
                shelt_wall['ShelterFactor'] = input_unit['shelterFactor'][id]
                shelt_wall['ShelterCode'] = None
                shelt_wall['Type'] = 'Cavity'
                shelt_wall['AreaCalculationType'] = 'Gross'
                shelt_wall['OpeningsArea'] = 'replace_xsi:nul'
                shelt_wall['NettArea'] = 0
                ext_walls['ExternalWall{}'.format(id)] = shelt_wall
            else:
                raise Exception(f'Warning for {input_unit["propertyName"]}: A sheltered wall has been entered without corresponding properties')
        elif type == 'Party wall':
            if row[6]>0:
                party_wall = {}
                party_wall['Description'] = input_unit['opaqElementName'][id]
                party_wall['Construction'] = 'Other'
                party_wall['Kappa'] = 0
                party_wall['GrossArea'] = round(input_unit['partylWallArea'][id],3)
                party_wall['Uvalue'] = 0
                party_wall['ShelterFactor'] = 0
                party_wall['ShelterCode'] = None
                party_wall['Type'] = 'FilledWithEdge'
                party_walls['PartyWall{}'.format(id)] = party_wall
            else:
                raise Exception('A party wall has been entered without corresponding properties')
        elif type == 'External roof':
            if row[7]>0 and row[8]>0 and row[9]!='nan' and row[10]>=0:
                ext_roof = {}
                ext_roof['Description'] = input_unit['opaqElementName'][id]
                ext_roof['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][id]-1))]
                ext_roof['Construction'] = 'Other'
                ext_roof['Kappa'] = 0
                ext_roof['GrossArea'] = input_unit['externalRoofArea'][id]
                ext_roof['Type'] = input_unit['externalRoofType'][id]
                ext_roof['UValue'] = input_unit['externalRoofUvalue'][id]
                ext_roof['ShelterFactor'] = input_unit['externalRoofShelterFactor'][id]
                ext_roof['ShelterCode'] = None
                ext_roof['AreaCalculationType'] = 'Gross'
                ext_roof['OpeningsArea'] = 'replace_xsi:nul'
                ext_roof['NettArea'] = 0
                ext_roofs['ExternalRoof{}'.format(id)] = ext_roof
            else:
                raise Exception(f'Warning for {input_unit["propertyName"]}: An external roof has been entered without corresponding properties')
        elif type == 'Heat loss floor':
            if row[11]>0 and row[12]>0 and row[13]!='nan' and row[14]>=0:
                heatloss_floor = {}
                heatloss_floor['Description'] = input_unit['opaqElementName'][id]
                heatloss_floor['Construction'] = 'Other'
                heatloss_floor['Kappa'] = 0
                heatloss_floor['Area'] = input_unit['heatLossFloorArea'][id]
                heatloss_floor['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][id]-1))]
                heatloss_floor['Type'] = input_unit['heatLossFloorType'][id]
                heatloss_floor['UValue'] = input_unit['heatLossFloorUvalue'][id]
                heatloss_floor['ShelterFactor'] = input_unit['heatLossFloorShelterFactor'][id]
                heatloss_floor['ShelterCode'] = None
                heatloss_floors['HeatLossFloor{}'.format(id)] = heatloss_floor
            else:
                raise Exception(f'Warning for {input_unit["propertyName"]}: A heat loss floor has been entered without corresponding properties')
        elif type == 'Party ceiling':
            if row[15]>0:
                party_roof = {}
                party_roof['Description'] = input_unit['opaqElementName'][id]
                party_roof['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][id]-1))]
                party_roof['Construction'] = 'Other'
                party_roof['Kappa'] = 0
                party_roof['GrossArea'] = input_unit['partyCeilingArea'][id]
                party_roofs['Roof{}'.format(id)] = party_roof
            else:
                raise Exception(f'Warning for {input_unit["propertyName"]}: A party ceiling has been entered without corresponding properties')
        elif type == 'Party floor':
            if row[16]>0:
                party_floor = {}
                party_floor['Description'] = input_unit['opaqElementName'][id]
                party_floor['Construction'] = 'Other'
                party_floor['Kappa'] = 0
                party_floor['Area'] = input_unit['partyFloorArea'][id]
                party_floor['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][id]-1))]
                party_floors['Floor{}'.format(id)] = party_floor
            else:
                raise Exception(f'Warning for {input_unit["propertyName"]}: A party floor has been entered without corresponding properties')
    
    # misc objects that need to be included in the XML for Elmhurst (but currently are not allowed to be entered in the excel sheet)
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
    
    # output data for lighting
    lightings = assessment['Lightings'] = {}
    lighting = lightings['Lighting'] = {}
    lighting['Name'] = input_unit['lightingName']
    lighting['Efficacy'] = input_unit['lightingEfficacy']
    lighting['Power'] = input_unit['lightingPower']
    lighting['Capacity'] = input_unit['lightingCapacity']
    lighting['Count'] = input_unit['lightingCount']
    
    # misc objects that need to be included in the XML for Elmhurst (but currently are not allowed to be entered in the excel sheet)
    assessment['ElectricityTariff'] ='Standard'
    assessment['SmartElectricityMeterFitted'] = 'false'
    assessment['SmartGasMeterFitted'] = 'false'
    assessment['SolarPanelPresent'] = 'false'
    assessment['PressureTest'] = 'true'
    assessment['PressureTestMethod'] = 'BlowerDoor'
    assessment['Designed_AP50_AP4'] = input_unit['airPerm']
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
    
    # output data for PV panels. Checks if PVs are present, otherwise retuns empty tag
    if input_unit['PVPresent'] == 'Yes':
        assessment['PhotovoltaicUnitType'] = input_unit['PVType']
        pvs = assessment['PhotovoltaicUnits'] = {}
        pv = pvs['PhotovoltaicUnit'] = {}
        pv['CellsPeak'] = input_unit['cellsPeak']
        pv['Orientation'] = input_unit['PVOrientation']
        pv['Elevation'] = input_unit['PVElevation']
        pv['Overshading'] = input_unit['PVOvershading']
        pv['Fghrs'] = 'false'
        pv['MCSCertificate'] = 'false'
        pv['OvershadingFactor'] = 0
    else:
        assessment['PhotovoltaicUnitType'] = None
        assessment['PhotovoltaicUnits'] = {}
    
    # output data for opening types
    assessment['OpeningTypes'] = {}
    opening_types = assessment['OpeningTypes'] = {}
    # looping through each opening type and only including if data is entered
    for id, name in enumerate(input_unit['openTypeName']):
        if input_unit['uVal'][id]>0:
            opening_type = {}
            opening_type['Description'] = name
            opening_type['DataSource'] = 'Manufacturer'
            opening_type['Type'] = input_unit['openingType'][id]
            if input_unit['openingType'][id]=='Window':
                opening_type['Glazing'] = 'Double'
                opening_type['GlazingGap'] = 'replace_xsi:nul'
                opening_type['GlazingFillingType'] = None
                opening_type['SolarTrans'] = input_unit['gVal'][id]
            else:
                opening_type['Glazing'] = 'replace_xsi:nul'
                opening_type['GlazingGap'] = 'replace_xsi:nul'
                opening_type['GlazingFillingType'] = None
                opening_type['SolarTrans'] = 0
            opening_type['FrameType'] = 'Wood'
            opening_type['FrameFactor'] = input_unit['frameFactor'][id]
            opening_type['UValue'] = input_unit['uVal'][id]
            opening_types['OpeningType{}'.format(id)] = opening_type
    
    # output data for openings
    openings = assessment['Openings'] = {}
    # looping through each opening and only including if data is entered
    for id, name in enumerate(input_unit['openName']):
        if input_unit['openArea'][id]>0:
            opening = {}
            opening['Id'] = id
            opening['OpeningTypeIndex'] = []
            for ido,type in enumerate(input_unit['openTypeName']):
                if input_unit['openType'][id] == type:
                    opening['OpeningTypeIndex'] = ido
            opening['Description'] = name
            opening['LocationBuildingPartIndex'] = 0
            for ido,parent in enumerate(input_unit['opaqElementName']):
                if input_unit['parentElem'][id] == parent:
                    opening['LocationWallIndex'] = ido
            opening['LocationRoofIndex'] = 'replace_xsi:nul'
            opening['Orientation'] = input_unit['openOrientation'][id]
            opening['AreaType'] = 'Total'
            opening['AreaScaleType'] = 'Meters'
            opening['Area'] = input_unit['openArea'][id]
            opening['AreaRecCalculation'] = []
            opening['RoofLightsPitch'] = 0
            openings['Opening{}'.format(id)] = opening
    
    # output data for thermal bridges
    thermal_bridges = assessment['ThermalBridges'] = {}
    # looping through each TB and only including if data is entered
    for tb,name in TBs.items():
        length = input_unit['thermalBridges'][tb]['lengths'][0]
        psi = input_unit['thermalBridges'][tb]['psi']
        if not math.isnan(length) and not math.isnan(psi):
            thermal_bridge = {}
            thermal_bridge['TypeSource'] = 'IndependentlyAssessed'
            thermal_bridge['Length'] = length
            thermal_bridge['PsiValue'] = psi
            thermal_bridge['K1Index'] = name
            thermal_bridge['Imported'] = 'False'
            thermal_bridge['Adjusted'] = psi
            thermal_bridge['Reference'] = []
            thermal_bridges['ThermalBridge{}'.format(tb)] = thermal_bridge
        # elif math.isnan(length) ^ math.isnan(psi):
        #     print(f'Warning for {input_unit["propertyName"]}: Thermal bridge {tb} has either a missing length or psi value')

    # output data for mech vent
    if input_unit['mechVentPresent'] == "Yes":
        mechvent = assessment['MechanicalVentilation'] = {}
        mechvent['DataType'] = input_unit['ventDataType']
        mechvent['Type'] = input_unit['ventType']
        mechvent['PcdfIndex'] = 'replace_xsi:nul'
        mechvent['PcdfItem'] = 'replace_xsi:nul'
        mechvent['ManufacturerSFP'] = input_unit['MVHR SFP']
        mechvent['DuctType'] = input_unit['ductType']
        mechvent['WetRooms'] = input_unit['wetRooms']
        mechvent['BrandModel'] = input_unit['brandModel']
        mechvent['MVHRDuctInsulated'] = 'replace_xsi:nul'
        mechvent['DuctInsulation'] = 'replace_xsi:nul'
        mechvent['MVHREfficiency'] = input_unit['MVHR HR'] 
        mechvent['ApprovedInstallation'] = "false"
        mechvent['SFPFromInstallerCertificate'] = "false"
        mechvent['MVHRSystemLocation'] = input_unit['systemLocation']
        if mechvent['MVHRSystemLocation'] == "Outside":
            mechvent['MVHRDuctInsulated'] = input_unit['ductInsulated']
        mechvent['DuctInsulationLevel'] = input_unit['ductInstallationSpecs']
    
    assessment['MechanicalVentilationDecentralised'] = []
    assessment['HeatingsInteraction'] = 'SeparatePartsOfHouse'
    
    # output data for the main heating systems. these are empty nested items that Elmhurst requires as input
    for mhs in range(2):
        mainHeatingSystem = assessment[f'MainHeatingSystem{mhs+1}'] = {}
        mainHeatingSystem['HeatingDataType'] = 'None'
        mainHeatingSystem['Fraction'] = 0
        mainHeatingSystem['PcdfIndex'] = 0
        mainHeatingSystem['BoilerEfficiencyType'] = 'replace_xsi:nul'
        mainHeatingSystem['EfficiencyWinter'] = 0
        mainHeatingSystem['EfficiencySummer'] = 0
        mainHeatingSystem['TestMethod'] = 'replace_xsi:nul'
        mainHeatingSystem['MHSCtrlPcdfIndex'] = 'replace_xsi:nul'
        mainHeatingSystem['CompensatorPcdfIndex'] = 'replace_xsi:nul'
        mainHeatingSystem['HetasApprovedSystem'] = 'false'
        mainHeatingSystem['FlueType'] = 'replace_xsi:nul'
        mainHeatingSystem['FanAssistedFlue'] = 'false'
        mainHeatingSystem['McsCertificate'] = 'false'
        mainHeatingSystem['Pumped'] = 'replace_xsi:nul'
        mainHeatingSystem['HeatingPumpAge'] = 'replace_xsi:nul'
        mainHeatingSystem['OilPumpInside'] = 'false'
        mainHeatingSystem['HeatEmitter'] = 'replace_xsi:nul'
        mainHeatingSystem['UnderfloorHeating'] = 'replace_xsi:nul'
        mainHeatingSystem['CombiType'] = 'replace_xsi:nul'
        mainHeatingSystem['CombiKeepHotType'] = 'replace_xsi:nul'
        mainHeatingSystem['CombiStoreType'] = 'replace_xsi:nul'
        mainHeatingSystem['ElectricCPSUtemperature'] = 'replace_xsi:nul'
        mainHeatingSystem['FIcase'] = 'replace_xsi:nul'
        mainHeatingSystem['FIwater'] = 'replace_xsi:nul'
        mainHeatingSystem['BurnerControl'] = 'replace_xsi:nul'
        mainHeatingSystem['DelayedStartStat'] = 'false'
        mainHeatingSystem['FlowTemperature'] = 'replace_xsi:nul'
        mainHeatingSystem['BoilerInterlock'] = 'false'
        mainHeatingSystem['StorageHeaters'] = {    }
        mainHeatingSystem['FlowTemperatureValue'] = 'replace_xsi:nul'
        mainHeatingSystem['SapCode'] = 'replace_xsi:nul'
        mainHeatingSystem['FuelType'] = 'replace_xsi:nul'
        mainHeatingSystem['CtrlSapCode'] = 'replace_xsi:nul'
    
    # output data for the secondary heating systems. this is an empty item that Elmhurst requires as input
    secondaryHeating = assessment['SecondaryHeating'] = {}
    secondaryHeating['HeatingDataType'] = 'None'
    secondaryHeating['TestMethod'] = 'replace_xsi:nul'
    secondaryHeating['HetasApprovedSystems'] = 'false'
    secondaryHeating['Efficiency'] = 'replace_xsi:nul'
    secondaryHeating['SapCode'] = 0
    secondaryHeating['FuelType'] = 'replace_xsi:nul'
    
    # output data for community heating
    communityHeating = assessment['CommunityHeating'] = {}
    communityHeating['Type'] = input_unit['heatingNetworkType']
    communityHeating['DistributionLossSpace'] = input_unit['distributionLossSpace']
    communityHeating['DistributionLossWater'] = 'replace_xsi:nul'
    communityHeating['ChargingLinked'] = 'replace_xsi:nul'
    heatSource = communityHeating['HeatSource'] = {}
    # Elmhurst expects 5 heat sources as input
    for chs in range(5):
        commHeatSource = heatSource[f'CommunityHeatSource{chs+1}'] = {}
        #only inputting into the first heat source, the rest is kept empty
        if chs == 0:
            commHeatSource['Source'] = input_unit['heatingSource1']
            commHeatSource['Fraction'] = input_unit['percentageHeat']
            commHeatSource['FuelType'] = input_unit['fuelType']
            commHeatSource['OveralEfficiency'] = input_unit['overallEfficiency']
            commHeatSource['HeatPowerRatio'] = 'replace_xsi:nul'
            commHeatSource['ElectricalEfficiency'] = 'replace_xsi:nul'
            commHeatSource['HeatEfficiency'] = 'replace_xsi:nul'
            commHeatSource['HeatingUse'] = input_unit['heatingUse']
            commHeatSource['CHPFuelFactor'] = 'replace_xsi:nul'
            commHeatSource['EfficiencyType'] = 'replace_xsi:nul'
        else:
            commHeatSource['Source'] = 'None'
            commHeatSource['Fraction'] = 'replace_xsi:nul'
            commHeatSource['FuelType'] = 'replace_xsi:nul'
            commHeatSource['OveralEfficiency'] = 'replace_xsi:nul'
            commHeatSource['HeatPowerRatio'] = 'replace_xsi:nul'
            commHeatSource['ElectricalEfficiency'] = 'replace_xsi:nul'
            commHeatSource['HeatEfficiency'] = 'replace_xsi:nul'
            commHeatSource['HeatingUse'] = input_unit['heatingUse']
            commHeatSource['CHPFuelFactor'] = 'replace_xsi:nul'
            commHeatSource['EfficiencyType'] = 'replace_xsi:nul'
            
    communityHeating['DistributionLossSpaceValue'] = input_unit['distributionLoss']
    communityHeating['DistributionLossWaterValue'] = 'replace_xsi:nul'
    communityHeating['SpacePCDFIndex'] = 'replace_xsi:nul'
    communityHeating['WaterPCDFIndex'] = 'replace_xsi:nul'
    communityHeating['CtrlSapCode'] = int(input_unit['heatingControls'])
    communityHeating['ExistingSpace'] = 'replace_xsi:nul'
    communityHeating['ExistingWater'] = 'replace_xsi:nul'
    communityHeating['UseNotionalSpace'] = 'replace_xsi:nul'
    communityHeating['UseNotionalWater'] = 'replace_xsi:nul'
    
    # output data for water heating systems DHW
    waterHeatingSystem = assessment['WaterHeatingSystem'] = {}
    waterHeatingSystem['WaterHeatingType'] = input_unit['waterHeating']
    waterHeatingSystem['LowWaterUse'] = 'false'
    waterHeatingSystem['ImmersionHeaterType'] = 'replace_xsi:nul'
    waterHeatingSystem['SummerImmersion'] = 'false'
    waterHeatingSystem['SuplementaryImmersion'] = 'false'
    waterHeatingSystem['ImmersionOnlyHeatingHotWater'] = 'false'
    waterHeatingSystem['ThermalStore'] = 'None'
    waterHeatingSystem['ThermalStorePipework'] = 'replace_xsi:nul'
    waterHeatingSystem['HotWaterCylinder'] = input_unit['storageType']
    waterHeatingSystem['InsulationType'] = 'replace_xsi:nul'
    waterHeatingSystem['InsulationThickness'] = 'replace_xsi:nul'
    waterHeatingSystem['InsulationThicknessType'] = 'replace_xsi:nul'
    waterHeatingSystem['Volume'] = 'replace_xsi:nul'
    waterHeatingSystem['CylinderStat'] = 'false'
    waterHeatingSystem['PipeworkInsulation'] = 'replace_xsi:nul'
    waterHeatingSystem['InHeatedSpace'] = 'false'
    waterHeatingSystem['InAiringCupboard'] = 'false'
    waterHeatingSystem['SeparateTimeControl'] = 0
    waterHeatingSystem['LossFactor'] = 1.46
    waterHeatingSystem['SolarPanelType'] = 'replace_xsi:nul'
    waterHeatingSystem['SolarAreaType'] = 'Aperture'
    waterHeatingSystem['SolarArea'] = 0
    waterHeatingSystem['SolarNi'] = 0
    waterHeatingSystem['SolarA1'] = 0
    waterHeatingSystem['SolarA2'] = 0
    waterHeatingSystem['SolarAGRatio'] = 0
    waterHeatingSystem['SolarLoopEfficiency'] = 0.9
    waterHeatingSystem['SolarKhem'] = 0
    waterHeatingSystem['SolarHeatLossCoeff'] = 'replace_xsi:nul'
    waterHeatingSystem['SolarIsFromCommunity'] = 'false'
    waterHeatingSystem['SolarServiceProvision'] = 'replace_xsi:nul'
    waterHeatingSystem['SolarPanelOrientation'] = 'replace_xsi:nul'
    waterHeatingSystem['SolarElevation'] = 'replace_xsi:nul'
    waterHeatingSystem['SolarOvershadingType'] = 'replace_xsi:nul'
    waterHeatingSystem['SolarVolume'] = 'replace_xsi:nul'
    waterHeatingSystem['SolarPumpElectricallyPowered'] = 'false'
    waterHeatingSystem['SolarCombinedCylinder'] = 'false'
    waterHeatingSystem['ColdWaterSource'] = input_unit['coldWaterSource']
    waterHeatingSystem['BathCount'] = int(input_unit['bathCount'])
    waterHeatingSystem['WWHRSBathCount'] = 'replace_xsi:nul'
    waterHeatingSystem['SapCode'] = 901
    waterHeatingSystem['FuelType'] = 'replace_xsi:nul'
    waterHeatingSystem['HIUPcdfIndex'] = 'replace_xsi:nul'
    
    # output data for showers (only one shower item is exported)
    showers = assessment['Showers'] = {}
    shower = showers['Shower'] = {}
    shower['Description'] = 'Shower1'
    shower['ShowerType'] = input_unit['showerType']
    shower['FlowRate'] = input_unit['showerFlowrate']
    shower['RatedPower'] = round(1.1625 * input_unit['showerFlowrate'],1)
    shower['Connected'] = 'true'
    shower['ConnectedTo'] = 'Storage'
    
    # misc objects that need to be included in the XML for Elmhurst (but currently are not allowed to be entered in the excel sheet)
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

def findAndReplace(file_path, value = 'replace_xsi:nul'):
    with open(file_path, 'r') as file:
        xml_data = file.read()
        root = ET.fromstring(xml_data)
        #add root attributes
        root.set("xmlns:xsd", "http://www.w3.org/2001/XMLSchema")
        root.set("xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance")
        #traverse XML and check if values match the flag string. If they do, remove it and add string as tag attribute
        def traverse(element):
            if element.text == value:
                old_tag = element.tag
                element.tag = '{} xsi:nil="true"'.format(old_tag)
                element.text = ''
            #check for integers in tags and remove them
            text = str(element.tag)
            for test_text in ['Measurement']:
                for i in range(10):
                    if text == test_text+str(i):
                        element.tag = test_text
            for test_text in ['ExternalWall','PartyWall','ExternalRoof','HeatLossFloor','Floor','Roof','OpeningType','Opening']:
                for i in range(20):
                    if text == test_text+str(i):
                        element.tag = test_text
            for test_text in ['CommunityHeatSource']:
                for i in range(5):
                    if text == test_text+str(i+1):
                        element.tag = test_text
            for tb,name in TBs.items():
                if text == 'ThermalBridge'+tb:
                    element.tag = 'ThermalBridge'
            for child in element:
                traverse(child)
        traverse(root)
        updated_xml_data = ET.tostring(root, encoding='utf-8', method='xml').decode()
    with open(file_path, 'w') as file:
        file.write(updated_xml_data)