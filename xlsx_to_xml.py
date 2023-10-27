from lxml import etree
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd


# class Result:

#     def __init__(self, basic_referenceValue, basic_measuredValue, basic_measurementError, uncertainty: float,
#                  distribution: str, k_factor: float, coverageProb: float, ds18b20_id: str):
#         d = {'basic_referenceValue': basic_referenceValue, 'basic_measuredValue': basic_measuredValue,
#              'basic_measurementError': basic_measurementError}

#         self.table = pd.DataFrame(data=d, index=[0, 1])
#         self.uncertainty = uncertainty
#         self.distribution = distribution
#         self.k_factor = k_factor
#         self.coverageProbability = coverageProb

#         self.ds18b20_id = ds18b20_id

class Result:

    def __init__(
                 self, No, modificationOrReposition, instrument, manufacturer, type, noSerie, 
                 noInternID, client, noRegister, certificateSheets, dateOfReception, dateOfCalibration, 
                 dateOfEmision, calibrate, authorized, atmPressure, pressureVariation, temperature, tempVariation, 
                 realtiveHumidity, humidityVariation, thermometerType, brand, model, indicationRange, escaleDivision, 
                 immersionIndicator, observations, conditionApplies, condition, calibrationMethod, methodHadAddition, 
                 describe, equipmentUsed, temperatureC, correction, uncertaintyExp, conformityDecApplies, decisionRule, 
                 techniquesOpinionApplies, techniquesOpinion, condition2, methodHadAddition2, conformityDec, techniquesOpinion2, 
                ):
        d = {'No': No, 'modificationOrReposition': modificationOrReposition, 'instrument': instrument, 'manufacturer': manufacturer,
             'type': type, 'noSerie': noSerie, 'noInternID': noInternID, 'client': client, 'noRegister': noRegister, 'certificateSheets': certificateSheets,
             'dateOfReception': dateOfReception, 'dateOfCalibration': dateOfCalibration, 'dateOfEmision': dateOfEmision, 'calibrate': calibrate,
             'authorized': authorized, 'atmPressure': atmPressure, 'pressureVariation': pressureVariation, 'temperature': temperature, 
             'tempVariation': tempVariation, 'realtiveHumidity': realtiveHumidity, 'humidityVariation': humidityVariation, 'thermometerType': thermometerType,
             'brand': brand, 'model': model, 'indicationRange': indicationRange, 'escaleDivision': escaleDivision, 'immersionIndicator': immersionIndicator,
             'observations': observations, 'conditionApplies': conditionApplies, 'condition': condition, 'calibrationMethod': calibrationMethod,
             'methodHadAddition': methodHadAddition, 'describe': describe, 'equipmentUsed': equipmentUsed, 'temperatureC': temperatureC,
             'correction': correction, 'uncertaintyExp': uncertaintyExp, 'conformityDecApplies': conformityDecApplies, 'decisionRule': decisionRule,
             'techniquesOpinionApplies': techniquesOpinionApplies, 'techniquesOpinion': techniquesOpinion, 'condition2': condition2,
             'methodHadAddition2': methodHadAddition2, 'conformityDec': conformityDec, 'techniquesOpinion2': techniquesOpinion2}

        self.table = pd.DataFrame(data=d, index=[0, 1])

class DCC:

    def __init__(self):

        self.dcc_original_data = None
        self.dcc_modified_data = None

    # loading DCC using etree.ElementTree
    def load_dcc(self, file):

        # parse xml
        try:
            ET.register_namespace('dcc', "https://cename.gt/dcc")
            ET.register_namespace('si', "https://cename.gt/si")
            doc = ET.parse(file)
            print('XML well formed, syntax ok.')
            self.dcc_original_data = doc
            return ET.tostring(self.dcc_original_data.getroot()).decode()

        # check for file IO error
        except IOError as err:
            print("IOError" + str(err))

        except ValueError as err:
            print("ValueError" + str(err))

        # check for XML syntax errors
        except etree.XMLSyntaxError as err:
            print('XML Syntax Error, see error_syntax.log')
            with open('error_syntax.log', 'w') as error_log_file:
                error_log_file.write(str(err))
            quit()

    def add_result_to_dcc(self, result: Result):
        xpath_basic_No = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_referenceValue\']/{https://cename.gt/si}hybrid/{https://cename.gt/si}realListXMLList[2]/{https://cename.gt/si}valueXMLList"
        xpath_basic_modificationOrReposition = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_referenceValue\']/{https://cename.gt/si}hybrid/{https://cename.gt/si}realListXMLList[1]/{https://cename.gt/si}valueXMLList"
        xpath_basic_instrument = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measuredValue\']/{https://cename.gt/si}hybrid/{https://cename.gt/si}realListXMLList[2]/{https://cename.gt/si}valueXMLList"
        xpath_basic_manufacturer = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measuredValue\']/{https://cename.gt/si}hybrid/{https://cename.gt/si}realListXMLList[1]/{https://cename.gt/si}valueXMLList"
        xpath_basic_type = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_noSerie = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_noInternID = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_client = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_noRegister = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_certificateSheets = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_dateOfReception = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_type = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_type = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_type = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"

        self.dcc_modified_data = self.dcc_original_data
        root = self.dcc_modified_data.getroot()

        basic_No = [round(element, 3) for element in result.table['basic_No'].tolist()]
        basic_modificationOrReposition = [round(element, 3) for element in result.table['basic_modificationOrReposition'].tolist()]
        basic_instrument = [round(element, 3) for element in result.table['basic_instrument'].tolist()]
        basic_manufacturer = [round(element, 3) for element in result.table['basic_manufacturer'].tolist()]
        basic_type = [round(element, 3) for element in result.table['basic_type'].tolist()]
        basic_noSerie = [round(element, 3) for element in result.table['basic_noSerie'].tolist()]
        basic_noInternID = [round(element, 3) for element in result.table['basic_noInternID'].tolist()]
        basic_client = [round(element, 3) for element in result.table['basic_client'].tolist()]
        basic_noRegister = [round(element, 3) for element in result.table['basic_noRegister'].tolist()]
        basic_certificateSheets = [round(element, 3) for element in result.table['basic_certificateSheets'].tolist()]
        basic_dateOfReception = [round(element, 3) for element in result.table['basic_dateOfReception'].tolist()]
        basic_No = [round(element, 3) for element in result.table['basic_No'].tolist()]

        try:
            root.find(xpath_basic_No).text = ' '.join(str(i) for i in basic_No)
        except AttributeError as e:
            print("root.find(xpath_basic_No).text" + str(e))

        try:
            root.find(xpath_basic_modificationOrReposition).text = ' '.join(str(i) for i in basic_modificationOrReposition)
        except AttributeError as e:
            print("root.find(xpath_basic_modificationOrReposition).text" + str(e))

        try:
            root.find(xpath_basic_instrument).text = ' '.join(str(i) for i in basic_instrument)
        except AttributeError as e:
            print("root.find(xpath_basic_instrument).text" + str(e))

        try:
            root.find(xpath_basic_manufacturer).text = ' '.join(str(i) for i in basic_manufacturer)
        except AttributeError as e:
            print("root.find(xpath_basic_manufacturer).text" + str(e))

        try:
            root.find(xpath_basic_type).text = ' '.join(str(i) for i in basic_type)
        except AttributeError as e:
            print("root.find(xpath_basic_type).text" + str(e))

        try:
            root.find(xpath_basic_noSerie).text = ' '.join(str(i) for i in basic_noSerie)
        except AttributeError as e:
            print("root.find(xpath_basic_noSerie).text" + str(e))

        try:
            root.find(xpath_basic_noInternID).text = ' '.join(str(i) for i in basic_noInternID)
        except AttributeError as e:
            print("root.find(xpath_basic_noInternID).text" + str(e))

        try:
            root.find(xpath_basic_client).text = ' '.join(str(i) for i in basic_client)
        except AttributeError as e:
            print("root.find(xpath_basic_client).text" + str(e))

        try:
            root.find(xpath_basic_noRegister).text = ' '.join(str(i) for i in basic_noRegister)
        except AttributeError as e:
            print("root.find(xpath_basic_noRegister).text" + str(e))

        try:
            root.find(xpath_basic_certificateSheets).text = ' '.join(str(i) for i in basic_certificateSheets)
        except AttributeError as e:
            print("root.find(xpath_basic_certificateSheets).text" + str(e))

        try:
            root.find(xpath_basic_dateOfReception).text = ' '.join(str(i) for i in basic_dateOfReception)
        except AttributeError as e:
            print("root.find(xpath_basic_dateOfReception).text" + str(e))

    def save_modified_dcc(self, path=None):
        ET.register_namespace('dcc', "https://cename.gt/dcc")
        ET.register_namespace('si', "https://cename.gt/si")
        file_timestamp = datetime.now().strftime("%Y_%m_%d-%I_%M_%S")
        filename = file_timestamp + " DCC modified.xml"
        self.dcc_modified_data.write(filename, encoding="utf8")

# if __name__ == "__main__":
#     mydcc = DCC()
#     mydcc.load_dcc("DCC_SummerSchool_unmodified.xml")
#     excel_data = pd.read_excel("data.xlsx")
#     my_results = Result(basic_referenceValue=excel_data['basic_referenceValue'].tolist(),
#                         basic_measuredValue=excel_data['basic_measuredValue'].tolist(),
#                         basic_measurementError=excel_data['basic_measurementError'].tolist(), 
#                         uncertainty=excel_data['uncertainty'].iloc[0],
#                         distribution=excel_data['distribution'].iloc[0],
#                         k_factor=excel_data['k_factor'].iloc[0],
#                         coverageProb=excel_data['coverageProb'].iloc[0],
#                         ds18b20_id=excel_data['ds18b20_id'].iloc[0]
#                         )

#     mydcc.add_result_to_dcc(my_results)
#     mydcc.save_modified_dcc(path="")

if __name__ == "__main__":
    mydcc = DCC()
    mydcc.load_dcc("DCC_SummerSchool_unmodified.xml")
    excel_data = pd.read_excel("data.xlsx")
    my_results = Result(
                            No=excel_data['No.'].tolist(),
                            modificationOrReposition=excel_data['Modificación o Reposición'].tolist(),
                            instrument=excel_data['Instrumento'].tolist(),
                            manufacturer=excel_data['Fabricante'].tolist(),
                            type=excel_data['Tipo'].tolist(),
                            noSerie=excel_data['No. de serie'].tolist(),
                            noInternID=excel_data['No. de identificación interna'].tolist(),
                            client=excel_data['Cliente'].tolist(),
                            noRegister=excel_data['No. de registro'].tolist(),
                            certificateSheets=excel_data['Hojas de este certificado'].tolist(),
                            dateOfReception=excel_data['Fecha de recepción'].tolist(),
                            dateOfCalibration=excel_data['Fecha de calibración'].tolist(),
                            dateOfEmision=excel_data['Fecha de emisión de certificado'].tolist(),
                            calibrate=excel_data['Calibró'].tolist(),
                            authorized=excel_data['Autorizó'].tolist(),
                            atmPressure=excel_data['Presión atmosférica'].tolist(),
                            pressureVariation=excel_data['Variación de presión'].tolist(),
                            temperature=excel_data['Temperatura'].tolist(),
                            tempVariation=excel_data['Varación de temperatura'].tolist(),
                            realtiveHumidity=excel_data['Humedad Relativa'].tolist(),
                            humidityVariation=excel_data['Variación de humedad'].tolist(),
                            thermometerType=excel_data['Tipo de termómetro'].tolist(),
                            brand=excel_data['Marca'].tolist(),
                            model=excel_data['Modelo'].tolist(),
                            indicationRange=excel_data['Intervalo de indicación'].tolist(),
                            escaleDivision=excel_data['División de escala'].tolist(),
                            immersionIndicator=excel_data['Indicador de profundidad de Inmersión'].tolist(),
                            observations=excel_data['Observaciones'].tolist(),
                            conditionApplies=excel_data['CONDICIÓN, aplica'].tolist(),
                            condition=excel_data['Condición'].tolist(),
                            calibrationMethod=excel_data['Método de calibración'].tolist(),
                            methodHadAddition=excel_data['El método utilizado por el CENAME tuvo alguna adición, desviación o exclusión al realizar la calibración del instrumento:'].tolist(),
                            describe=excel_data['Describa:'].tolist(),
                            equipmentUsed=excel_data['Patrones y equipo utilizados'].tolist(),
                            temperatureC=excel_data['Temperatura °C'].tolist(),
                            correction=excel_data['Corrección °C'].tolist(),
                            uncertaintyExp=excel_data['Incertidumbre Expandida ± °C'].tolist(),
                            conformityDecApplies=excel_data['DECLARACIÓN DE LA CONFORMIDAD (de ser aplicable)'].tolist(),
                            decisionRule=excel_data['Describa Regla de Decisión'].tolist(),
                            techniquesOpinionApplies=excel_data['OPINIONES/INTERPRETACIONES TÉCNICAS (de ser aplicable)'].tolist(),
                            techniquesOpinion=excel_data['Opiniones/Interpretaciones Técnicas'].tolist(),
                            condition2=excel_data['Condición'].tolist(),
                            methodHadAddition2=excel_data['El método utilizado por el CENAME tuvo alguna adición, desviación o exclusión al realizar la calibración del instrumento:'].tolist(),
                            conformityDec=excel_data['Declaración de la conformidad'].tolist(),
                            techniquesOpinion2=excel_data['Opiniones/interpretaciones técnicas'].tolist()
                            
                        )

    mydcc.add_result_to_dcc(my_results)
    mydcc.save_modified_dcc(path="")

# if __name__ == "__main__":
#     mydcc = DCC()
#     mydcc.load_dcc("DCC_SummerSchool_unmodified.xml")
#     excel_data = pd.read_excel("data.xlsx", names = [
#         "No.", "Modificación o Reposición", "Instrumento", "Fabricante", "Tipo", "No. de serie",
#         "No. de identificación interna", "Cliente", "No. de registro", "Hojas de este certificado",
#         "Fecha de recepción", "Fecha de calibración", "Fecha de emisión de certificado", "Calibró",
#         "Autorizó", "Presión atmosférica", "Variación de presión", "Temperatura", "Variación de temperatura",
#         "Humedad Relativa", "Variación de humedad", "Tipo de termómetro", "Marca", "Modelo",
#         "Intervalo de indicación", "División de escala", "Indicador de profundidad de Inmersión",
#         "Observaciones", "CONDICIÓN, aplica", "Condición", "Método de calibración",
#         "El método utilizado por el CENAME tuvo alguna adición, desviación o exclusión al realizar la calibración del instrumento",
#         "Describa", "Patrones y equipo utilizados", "Temperatura °C", "Corrección °C",
#         "Incertidumbre Expandida ± °C", "Temperatura °F", "Corrección °F", "Incertidumbre Expandida ± °F",
#         "DECLARACIÓN DE LA CONFORMIDAD (de ser aplicable)", "Describa Regla de Decisión",
#         "OPINIONES/INTERPRETACIONES TÉCNICAS (de ser aplicable)", "Opiniones/Interpretaciones Técnicas",
#         "Condición", "El método utilizado por el CENAME tuvo alguna adición, desviación o exclusión al realizar la calibración del instrumento",
#         "Declaración de la conformidad", "Opiniones/interpretaciones técnicas"
#     ])

#     my_results = Result(
#         basic_referenceValue=excel_data['Temperatura °C'].tolist(),
#         basic_measuredValue=excel_data['Corrección °C'].tolist(),
#         basic_measurementError=excel_data['Incertidumbre Expandida ± °C'].tolist(), 
#         uncertainty=excel_data['Temperatura °F'].iloc[0],
#         distribution=excel_data['Corrección °F'].iloc[0],
#         k_factor=excel_data['Incertidumbre Expandida ± °F'].iloc[0],
#         coverageProb=excel_data['DECLARACIÓN DE LA CONFORMIDAD (de ser aplicable)'].iloc[0],
#         ds18b20_id=excel_data['Describa Regla de Decisión'].iloc[0]
#     )

#     mydcc.add_result_to_dcc(my_results)
#     mydcc.save_modified_dcc(path="")


    



