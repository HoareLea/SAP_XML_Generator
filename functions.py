import xml.etree.ElementTree as ET
from xml.dom import minidom
from TBs import TBs
from levels_naming import levels_naming
import math

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

def input_reader(sheet):
    unit = {}
    
    # General information
    unit['propertyName'] = sheet['Property name'].iloc[1]
    unit['dwellingOrientation'] = sheet['Dwelling orientation'].iloc[1]
    unit['calcType'] = sheet['Calculation type'].iloc[1]
    # unit['propertyTenure'] = sheet['Property tenure'].iloc[1]
    # unit['transactionType'] = sheet['Transaction type'].iloc[1]
    unit['terrainType'] = sheet['Terrain type'].iloc[1]
    unit['propertyType1'] = sheet['Property type 1'].iloc[1]
    unit['propertyType2'] = sheet['Property type 2'].iloc[1]
    unit['positionOfFlat'] = sheet['Position of flat'].iloc[1]
    unit['whichFloor'] = str(int(sheet['Which floor'].iloc[1]))
    unit['storeysInBlock'] = str(int(sheet['Tot no. storeys in block'].iloc[1]))
    unit['noOfStoreys'] = sheet['No. storeys'].iloc[1]
    unit['dateBuilt'] = sheet['Date built'].iloc[1]
    unit['shelteredSides'] = sheet['Sheltered sides'].iloc[1]
    unit['sunlightSunshade'] = sheet['Sunlight/sunshade'].iloc[1]
    unit['tmp'] = sheet['Thermal mass parameter'].iloc[1]
    unit['livingArea'] = sheet['Living area'].iloc[1]
    
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
    unit['shelterFactor'] = sheet['Shelter factor'].tolist()[1:]
    unit['partylWallArea'] = sheet['Party wall area'].tolist()[1:]
    unit['roofArea'] = sheet['Roof area'].tolist()[1:]
    unit['roofUvalue'] = sheet['Roof U-value'].tolist()[1:]
    unit['partyRoofArea'] = sheet['Party roof area'].tolist()[1:]
    unit['heatLossFloorArea'] = sheet['Heat loss floor area'].tolist()[1:]
    unit['heatLossFloorUvalue'] = sheet['Heat loss floor U-value'].tolist()[1:]
    unit['partyFloorArea'] = sheet['Party floor area'].tolist()[1:]
    unit['groundContArea'] = sheet['Ground-contact floor area'].tolist()[1:]
    unit['groundContUvalue'] = sheet['Ground-contact floor U-value'].tolist()[1:]

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
    for tb,name in TBs.items():
        thermal_bridge = {}
        thermal_bridge['psi'] = sheet[tb].tolist()[0]
        thermal_bridge['lengths']  = sheet[tb].tolist()[2:]
        thermal_bridges[tb] = thermal_bridge
    return unit

def match_xml(input_unit):
    output_data = {}
    assessment = output_data['Assessment'] = {}
    assessment['Reference'] = input_unit['propertyName']
    assessment['DwellingOrientation'] = input_unit['dwellingOrientation']
    assessment['CalculationType'] = input_unit['calcType']
    assessment['Tenure'] = 'ND' #input_unit['propertyTenure']
    assessment['TransactionType'] = str(int(6)) #input_unit['transactionType']
    assessment['TerrainType'] = input_unit['terrainType']
    assessment['SimpleComplianceScotland'] = 'false'
    assessment['PropertyType1'] = input_unit['propertyType1']
    assessment['PropertyType2'] = input_unit['propertyType2']
    assessment['PositionOfFlat'] = input_unit['positionOfFlat']
    assessment['FlatWhichFloor'] = input_unit['whichFloor']
    assessment['StoreysInBlock'] = input_unit['storeysInBlock']
    assessment['Storeys'] = sum(1 for x in input_unit['floorToSlab'] if x > 0)
    assessment['DateBuilt'] = str(int(input_unit['dateBuilt']))
    assessment['PropertyAgeBand'] = 'replace_xsi:nul'
    assessment['ShelteredSides'] = str(int(input_unit['shelteredSides']))
    assessment['SunlightShade'] = input_unit['sunlightSunshade']
    assessment['Basement'] = 'false'
    assessment['LivingArea'] = input_unit['livingArea']
    assessment['ThermalMass'] = 'EnterTmpValue'
    assessment['ThermalMassValue'] = input_unit['tmp']
    assessment['LowestFloorHasUnheatedSpace'] = 'replace_xsi:nul'
    assessment['UnheatedFloorArea'] = 'replace_xsi:nul'
    measurements = assessment['Measurements'] = {}
    for msrmt in range(9):
        measurement = {}
        msrmt_len = len(input_unit['heatLossPerim'])
        if msrmt>0 and msrmt<msrmt_len:
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
        else:
            measurement['Storey'] = msrmt
            measurement['InternalPerimeter'] = 0
            measurement['InternalFloorArea'] = 0
            measurement['StoreyHeight'] = 0
        measurements['Measurement{}'.format(msrmt)]  = measurement
    ext_walls = assessment['ExternalWalls'] = {}
    party_walls = assessment['PartyWalls'] = {}
    assessment['InternalPartitions'] = []
    ext_roofs = assessment['ExternalRoofs'] = {}
    party_roofs = assessment['PartyRoofs'] = {}
    assessment['InternalCeilings'] = []
    heatloss_floors = assessment['HeatlossFloors'] = {}
    party_floors = assessment['PartyFloors'] = {}
    assessment['InternalFloors'] = []
    for id, type in enumerate(input_unit['opaqElementType']):
        if type == 'External wall':
            ext_wall = {}
            ext_wall['Description'] = input_unit['opaqElementName'][id]
            ext_wall['Construction'] = 'Other'
            ext_wall['Kappa'] = 0
            ext_wall['GrossArea'] = input_unit['externalWallArea'][id]
            ext_wall['Uvalue'] = input_unit['externalWallUvalue'][id]
            ext_wall['ShelterFactor'] = 0
            ext_wall['ShelterCode'] = None
            ext_wall['Type'] = 'Cavity'
            ext_wall['AreaCalculationType'] = 'Gross'
            ext_wall['OpeningsArea'] = 'replace_xsi:nul'
            ext_wall['NettArea'] = 0
            ext_walls['ExternalWall{}'.format(id)] = ext_wall
        elif type == 'Sheltered wall':
            shelt_wall = {}
            shelt_wall['Description'] = input_unit['opaqElementName'][id]
            shelt_wall['Construction'] = 'Other'
            shelt_wall['Kappa'] = 0
            shelt_wall['GrossArea'] = input_unit['shelteredWallArea'][id]
            shelt_wall['Uvalue'] = input_unit['shelteredWallUvalue'][id]
            shelt_wall['ShelterFactor'] = input_unit['shelterFactor'][id]
            shelt_wall['ShelterCode'] = None
            shelt_wall['Type'] = 'Cavity'
            shelt_wall['AreaCalculationType'] = 'Gross'
            shelt_wall['OpeningsArea'] = 'replace_xsi:nul'
            shelt_wall['NettArea'] = 0
            ext_walls['ExternalWall{}'.format(id)] = shelt_wall
        elif type == 'Party wall':
            party_wall = {}
            party_wall['Description'] = input_unit['opaqElementName'][id]
            party_wall['Construction'] = 'Other'
            party_wall['Kappa'] = 0
            party_wall['GrossArea'] = input_unit['partylWallArea'][id]
            party_wall['Uvalue'] = 0
            party_wall['ShelterFactor'] = 0
            party_wall['ShelterCode'] = None
            party_wall['Type'] = 'FilledWithEdge'
            party_walls['PartyWall{}'.format(id)] = party_wall
        elif type == 'External roof':
            ext_roof = {}
            ext_roof['Description'] = input_unit['opaqElementName'][id]
            ext_roof['StoreyIndex'] = []
            ext_roof['Construction'] = 'Other'
            ext_roof['Kappa'] = 0
            ext_roof['GrossArea'] = input_unit['externalRoofArea'][id]
            ext_roof['Type'] = input_unit['roofType'][id]
            ext_roof['Uvalue'] = input_unit['externalRoofUvalue'][id]
            ext_roof['ShelterFactor'] = 0
            ext_roof['ShelterCode'] = None
            ext_roof['AreaCalculationType'] = 'Gross'
            ext_roof['OpeningsArea'] = []
            ext_roof['NettArea'] = []
            ext_roofs['ExternalRoof{}'.format(id)] = ext_roof
        elif type == 'Heat loss floor':
            heatloss_floor = {}
            heatloss_floor['Description'] = input_unit['opaqElementName'][id]
            heatloss_floor['Construction'] = 'Other'
            heatloss_floor['Kappa'] = 0
            heatloss_floor['Area'] = input_unit['heatLossFloorArea'][id]
            heatloss_floor['StoreyIndex'] = []
            heatloss_floor['Type'] = 'Cavity'
            heatloss_floor['Uvalue'] = input_unit['heatLossFloorUvalue'][id]
            heatloss_floor['ShelterFactor'] = 0
            heatloss_floor['ShelterCode'] = None
            heatloss_floors['HeatLossFloor{}'.format(id)] = heatloss_floor
        elif type == 'Party floor':
            party_floor = {}
            party_floor['Description'] = input_unit['opaqElementName'][id]
            party_floor['Construction'] = 'Other'
            party_floor['Kappa'] = 0
            party_floor['Area'] = input_unit['partyFloorArea'][id]
            party_floor['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][id]-1))]
            party_floors['Floor{}'.format(id)] = party_floor
        elif type == 'Party ceiling':
            party_roof = {}
            party_roof['Description'] = input_unit['opaqElementName'][id]
            party_roof['StoreyIndex'] = levels_naming[str(int(input_unit['opaqElementLevel'][id]-1))]
            party_roof['Construction'] = 'Other'
            party_roof['Kappa'] = 0
            party_roof['GrossArea'] = input_unit['partyRoofArea'][id]
            party_roofs['Roof{}'.format(id)] = party_roof
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
    assessment['Lightings'] = {}
    assessment['ElectricityTariff'] ='Standard'
    assessment['SmartElectricityMeterFitted'] = 'false'
    assessment['SmartGasMeterFitted'] = 'false'
    assessment['SolarPanelPresent'] = 'false'
    assessment['PressureTest'] = 'true'
    assessment['PressureTestMethod'] = 'BlowerDoor'
    assessment['Designed_AP50_AP4'] = 3.0
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
    assessment['BatteryCapacity'] = 'replace_xsi:nul'
    assessment['PhotovoltaicUnitType'] = None
    assessment['PhotovoltaicUnits'] = []
    assessment['OpeningTypes'] = {}
    opening_types = assessment['OpeningTypes'] = {}
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
    openings = assessment['Openings'] = {}
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
    thermal_bridges = assessment['ThermalBridges'] = {}
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
    assessment['MechanicalVentilation'] = {}
    assessment['MechanicalVentilationDecentralised'] = []
    assessment['HeatingsInteraction'] = 'SeparatePartsOfHouse'
    assessment['CommunityHeating'] = {}
    assessment['WaterHeatingSystem'] = {}
    assessment['Showers'] = {}
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
            for test_text in ['ExternalWall','PartyWall','ExternalRoof','HeatlossFloor','Floor','Roof','OpeningType','Opening']:
                for i in range(20):
                    if text == test_text+str(i):
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