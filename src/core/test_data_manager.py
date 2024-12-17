from typing import Dict, List
from data_processor import TestData, calc_NR_new, calc_atl_val, calc_AIIC_val_claude, calc_astc_val

class TestDataManager:
    def __init__(self):
        self.tests: Dict[str, TestData] = {}
    
    def get_test_list(self) -> List[str]:
        """Return sorted list of available test IDs by test_label"""
        # Sort by test_label from the existing room_properties
        return sorted(
            self.tests.keys(),
            key=lambda x: self.tests[x].room_properties['Test_Label']
        )
    
    def get_test_data(self, test_id: str) -> TestData:
        """Return TestData object for specified test"""
        return self.tests.get(test_id)
    
    def load_test_data(self, test_id: str, raw_data: dict) -> None:
        """Process raw data using existing calculation functions"""
        try:
            # Use existing calculation functions from data_processor.py
            NR_val, NIC_final_val, sabines, AIIC_recieve_corr, ASTC_recieve_corr, AIIC_Normalized_recieve = calc_NR_new(
                raw_data['srs_overalloct'],
                raw_data['AIIC_rec_overalloct'],
                raw_data['ASTC_rec_overalloct'],
                raw_data['bkgrnd_overalloct'],
                raw_data['bkgrnd_overalloct_aiic'],
                raw_data['rt_thirty'],
                raw_data['STCCurve'],
                raw_data['NIC_start']
            )
            
            # Store processed data in TestData object
            test_data = TestData()
            test_data.room_properties = raw_data['room_properties']
            test_data.NR_val = NR_val
            test_data.NIC_final_val = NIC_final_val
            test_data.sabines = sabines
            test_data.AIIC_recieve_corr = AIIC_recieve_corr
            test_data.ASTC_recieve_corr = ASTC_recieve_corr
            test_data.AIIC_Normalized_recieve = AIIC_Normalized_recieve
            
            # Store using test_label as key
            self.tests[test_id] = test_data
            
        except Exception as e:
            print(f"Error processing test data: {str(e)}")
            raise

    def get_sorted_tests(self) -> List[TestData]:
        """Return list of TestData objects sorted by test_label"""
        return [
            self.tests[test_id] 
            for test_id in self.get_test_list()
        ] 