from lxml import etree
import xml.etree.ElementTree as ET
from datetime import datetime
import pandas as pd


class Result:

    def __init__(self, basic_referenceValue, basic_measuredValue, basic_measurementError, uncertainty: float,
                 distribution: str, k_factor: float, coverageProb: float, ds18b20_id: str):
        d = {'basic_referenceValue': basic_referenceValue, 'basic_measuredValue': basic_measuredValue,
             'basic_measurementError': basic_measurementError}

        self.table = pd.DataFrame(data=d, index=[0, 1])
        self.uncertainty = uncertainty
        self.distribution = distribution
        self.k_factor = k_factor
        self.coverageProbability = coverageProb

        self.ds18b20_id = ds18b20_id


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
        xpath_basic_referenceValue_C = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_referenceValue\']/{https://cename.gt/si}hybrid/{https://cename.gt/si}realListXMLList[2]/{https://cename.gt/si}valueXMLList"
        xpath_basic_referenceValue_K = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_referenceValue\']/{https://cename.gt/si}hybrid/{https://cename.gt/si}realListXMLList[1]/{https://cename.gt/si}valueXMLList"
        xpath_basic_measuredValue_C = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measuredValue\']/{https://cename.gt/si}hybrid/{https://cename.gt/si}realListXMLList[2]/{https://cename.gt/si}valueXMLList"
        xpath_basic_measuredValue_K = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measuredValue\']/{https://cename.gt/si}hybrid/{https://cename.gt/si}realListXMLList[1]/{https://cename.gt/si}valueXMLList"
        xpath_basic_measurementError_K = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}valueXMLList"
        xpath_basic_measurementUncertainty = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}expandedUncXMLList/{https://cename.gt/si}uncertaintyXMLList"
        xpath_basic_coverageFactor = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}expandedUncXMLList/{https://cename.gt/si}coverageFactorXMLList"
        xpath_basic_coverageProbability = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}expandedUncXMLList/{https://cename.gt/si}coverageProbabilityXMLList"
        xpath_basic_distribution = ".//{https://cename.gt/dcc}quantity[@refType=\'basic_measurementError\']/{https://cename.gt/si}realListXMLList/{https://cename.gt/si}expandedUncXMLList/{https://cename.gt/si}distributionXMLList"

        xpath_item_ID = ".//{https://cename.gt/dcc}item[2]/{https://cename.gt/dcc}identifications/{https://cename.gt/dcc}identification[{https://cename.gt/dcc}issuer=\'manufacturer\']/{https://cename.gt/dcc}value"

        self.dcc_modified_data = self.dcc_original_data
        root = self.dcc_modified_data.getroot()

        basic_refValue_C = [round(element, 3) for element in result.table['basic_referenceValue'].tolist()]

        basic_refValue_K = [round(element, 3) for element in
                            (result.table['basic_referenceValue'] + 273.15).tolist()]

        basic_measuredValue_C = [round(element, 3) for element in result.table['basic_measuredValue'].tolist()]

        basic_measuredValue_K = [round(element, 3) for element in
                                 (result.table['basic_measuredValue'] + 273.15).tolist()]

        basic_measurementError_K = [round(element, 3) for element in
                                    result.table['basic_measurementError'].tolist()]

        try:
            # set basic_referenceValue in celcius
            root.find(xpath_basic_referenceValue_C).text = ' '.join(str(i) for i in basic_refValue_C)
            # set basic_referenceValue in kelvin
        except AttributeError as e:
            print("root.find(xpath_basic_referenceValue_C).text" + str(e))

        try:
            root.find(xpath_basic_referenceValue_K).text = ' '.join(str(i) for i in basic_refValue_K)
        except AttributeError as e:
            print("root.find(xpath_basic_referenceValue_K).text" + str(e))

        try:
            # set basic_measuredValue in celcius
            root.find(xpath_basic_measuredValue_C).text = ' '.join(str(i) for i in basic_measuredValue_C)
        except AttributeError as e:
            print("root.find(xpath_basic_measuredValue_C).text" + str(e))

        try:
            # set basic_measuredValue in kelvin
            root.find(xpath_basic_measuredValue_K).text = ' '.join(str(i) for i in basic_measuredValue_K)
        except AttributeError as e:
            print("root.find(xpath_basic_measuredValue_K).text" + str(e))

        try:
            # set basic_measurementError
            root.find(xpath_basic_measurementError_K).text = ' '.join(str(i) for i in basic_measurementError_K)
        except AttributeError as e:
            print("root.find(xpath_basic_measurementError_K).text" + str(e))

        try:
            # set measurementUncertainty, k factor, coverage probability and distribution
            root.find(xpath_basic_measurementUncertainty).text = str(result.uncertainty)
        except AttributeError as e:
            print("root.find(xpath_basic_measurementUncertainty).text" + str(e))

        try:
            root.find(xpath_basic_coverageFactor).text = str(result.k_factor)
        except AttributeError as e:
            print("root.find(xpath_basic_coverageFactor).text" + str(e))

        try:
            root.find(xpath_basic_coverageProbability).text = str(result.coverageProbability)
        except AttributeError as e:
            print("root.find(xpath_basic_coverageProbability).text" + str(e))

        try:
            root.find(xpath_basic_distribution).text = str(result.distribution)
        except AttributeError as e:
            print("root.find(xpath_basic_distribution).text" + str(e))

        try:
            root.find(xpath_item_ID).text = result.ds18b20_id
        except AttributeError as e:
            print("root.find(xpath_item_ID).text" + str(e))

    def save_modified_dcc(self, path=None):
        ET.register_namespace('dcc', "https://cename.gt/dcc")
        ET.register_namespace('si', "https://cename.gt/si")
        file_timestamp = datetime.now().strftime("%Y_%m_%d-%I_%M_%S")
        filename = file_timestamp + " DCC modified.xml"
        self.dcc_modified_data.write(filename, encoding="utf8")



if __name__ == "__main__":
    mydcc = DCC()
    mydcc.load_dcc("DCC_SummerSchool_unmodified.xml")
    excel_data = pd.read_excel("C:/Users/ME-02/OneDrive/Documentos/Plan CABUREK/XML DCC/De excel a xml en python/lxml")
    my_results = Result(basic_referenceValue=excel_data['basic_referenceValue'].tolist(),
                        basic_measuredValue=excel_data['basic_measuredValue'].tolist(),
                        basic_measurementError=excel_data['basic_measurementError'].tolist(), 
                        uncertainty=excel_data['uncertainty'].iloc[0],
                        distribution=excel_data['distribution'].iloc[0],
                        k_factor=excel_data['k_factor'].iloc[0],
                        coverageProb=excel_data['coverageProb'].iloc[0],
                        ds18b20_id=excel_data['ds18b20_id'].iloc[0]
                        )



    mydcc.add_result_to_dcc(my_results)
    mydcc.save_modified_dcc(path="")
    



