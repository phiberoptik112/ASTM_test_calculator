from typing import Dict, Optional, List, Union
import pandas as pd
from pathlib import Path

from data_processor import (
    TestType,
    RoomProperties,
    AIICTestData,
    ASTCTestData,
    NICTestData,
    DTCtestData
)
from src.core.test_processor import TestProcessor

class TestDataManager:
    def __init__(self, debug_mode: bool = False):
        self.debug_mode = debug_mode
        self.test_processor = TestProcessor(debug_mode=debug_mode)
        self.test_data_collection: Dict[str, Dict[TestType, Dict]] = {}
        self.test_plan: Optional[pd.DataFrame] = None
        self.slm_data_paths: Dict[str, str] = {
            'meter_1': '',
            'meter_2': ''
        }
        
    def set_data_paths(self, meter_1_path: str, meter_2_path: str) -> None:
        """Set paths for SLM data files"""
        self.slm_data_paths['meter_1'] = meter_1_path
        self.slm_data_paths['meter_2'] = meter_2_path

    def load_test_plan(self, test_plan_path: str) -> None:
        """Load test plan from Excel file"""
        try:
            self.test_plan = pd.read_excel(test_plan_path)
            if self.debug_mode:
                print(f"Loaded test plan with {len(self.test_plan)} rows")
                print("Columns:", self.test_plan.columns.tolist())
        except Exception as e:
            raise ValueError(f"Error loading test plan: {str(e)}")

    def process_test_data(self) -> None:
        """Process all tests in the test plan"""
        if self.test_plan is None:
            raise ValueError("No test plan loaded")

        for index, curr_test in self.test_plan.iterrows():
            try:
                test_label = curr_test['Test_Label']
                if self.debug_mode:
                    print(f"\nProcessing test {test_label}")

                # Create room properties
                room_props = self._create_room_properties(curr_test)
                
                # Process each test type
                curr_test_data = {}
                for test_type, column in {
                    TestType.AIIC: 'AIIC',
                    TestType.ASTC: 'ASTC',
                    TestType.NIC: 'NIC',
                    TestType.DTC: 'DTC'
                }.items():
                    if curr_test[column] == 1:
                        try:
                            test_data = self.test_processor.load_test_data(
                                curr_test=curr_test,
                                test_type=test_type,
                                room_props=room_props,
                                slm_data_paths=self.slm_data_paths
                            )
                            curr_test_data[test_type] = {
                                'room_properties': room_props,
                                'test_data': test_data
                            }
                        except Exception as e:
                            if self.debug_mode:
                                print(f"Error processing {test_type.value} test: {str(e)}")

                if curr_test_data:
                    self.test_data_collection[test_label] = curr_test_data

            except Exception as e:
                if self.debug_mode:
                    print(f"Error processing test row {index}: {str(e)}")

    def get_test_data(self, test_label: str, test_type: TestType) -> Optional[Dict]:
        """Retrieve test data for specific test and type"""
        try:
            return self.test_data_collection[test_label][test_type]
        except KeyError:
            return None

    def get_all_test_labels(self) -> List[str]:
        """Get list of all test labels"""
        return list(self.test_data_collection.keys())

    def get_test_types_for_label(self, test_label: str) -> List[TestType]:
        """Get list of test types available for a test label"""
        try:
            return list(self.test_data_collection[test_label].keys())
        except KeyError:
            return []

    def _create_room_properties(self, test_row: pd.Series) -> RoomProperties:
        """Create RoomProperties instance from test row"""
        return RoomProperties(
            site_name=test_row['Site_Name'],
            client_name=test_row['Client_Name'],
            source_room=test_row['Source Room'],
            receive_room=test_row['Receiving Room'],
            test_date=test_row['Test Date'],
            report_date=test_row['Report Date'],
            project_name=test_row['Project Name'],
            test_label=test_row['Test_Label'],
            source_vol=test_row['source room vol'],
            receive_vol=test_row['receive room vol'],
            partition_area=test_row['partition area'],
            partition_dim=test_row['partition dim'],
            source_room_finish=test_row['source room finish'],
            receive_room_finish=test_row['receive room finish'],
            srs_floor=test_row['srs floor descrip.'],
            srs_walls=test_row['srs walls descrip.'],
            srs_ceiling=test_row['srs ceiling descrip.'],
            rec_floor=test_row['rec floor descrip.'],
            rec_walls=test_row['rec walls descrip.'],
            rec_ceiling=test_row['rec ceiling descrip.'],
            tested_assembly=test_row['tested assembly'],
            expected_performance=test_row['expected performance'],
            test_assembly_type=test_row['Test assembly Type'],
            annex_2_used=bool(test_row['Annex 2 used?'])
        )