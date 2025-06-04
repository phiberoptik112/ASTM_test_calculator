import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
from typing import Dict, List, Union
from src.core.data_processor import TestType
import logging
import sys

# Set up logging with more detailed format
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    stream=sys.stdout
)
logger = logging.getLogger(__name__)

def plot_curves(frequencies: List[float], y_label: str, ref_curve: np.ndarray, 
                field_curve: np.ndarray, ref_label: str, field_label: str) -> str:
    """Generate a plot comparing reference and field curves"""
    plt.figure(figsize=(10, 6))
    plt.semilogx(frequencies, ref_curve, 'k--', label=ref_label)
    plt.semilogx(frequencies, field_curve, 'k-', label=field_label)
    plt.grid(True, which="both", ls="-", alpha=0.2)
    plt.xlabel('Frequency (Hz)')
    plt.ylabel(y_label)
    plt.legend()
    
    # Save plot to temporary file
    plot_path = Path('temp_plot.png')
    plt.savefig(plot_path)
    plt.close()
    
    return str(plot_path)

def export_test_results(test_data_collection: Dict, output_dir: Union[str, Path]) -> None:
    """Export test results to CSV and generate plots
    
    Args:
        test_data_collection: Dictionary containing test results
        output_dir: Directory to save output files
    """
    logger.debug(f"Received test_data_collection: {test_data_collection}")
    
    output_dir = Path(output_dir).resolve()  # Get absolute path
    logger.info(f"Creating output directory at: {output_dir}")
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Create DataFrames for each test type
    aiic_data = []
    astc_data = []
    nic_data = []
    
    # Process each test
    for test_label, test_types in test_data_collection.items():
        logger.info(f"Processing test: {test_label}")
        logger.debug(f"Test types for {test_label}: {test_types}")
        
        for test_type, test_info in test_types.items():
            logger.debug(f"Processing test type {test_type} for {test_label}")
            logger.debug(f"Test info: {test_info}")
            
            test_data = test_info['test_data']
            logger.debug(f"Test data type: {type(test_data)}")
            logger.debug(f"Test data attributes: {dir(test_data)}")
            
            if not hasattr(test_data, 'calculated_values'):
                logger.error(f"No calculated_values found in test_data for {test_label} {test_type}")
                continue
                
            calc_values = test_data.calculated_values
            logger.debug(f"Calculated values for {test_label} {test_type}: {calc_values}")
            
            if test_type == TestType.AIIC:
                logger.info(f"Processing AIIC data for {test_label}")
                try:
                    # Convert numpy arrays to lists for JSON serialization
                    aiic_entry = {
                        'Test_Label': test_label,
                        'NR_val': calc_values['NR_val'] if calc_values['NR_val'] is not None else None,
                        'AIIC_recieve_corr': calc_values['AIIC_recieve_corr'].tolist() if hasattr(calc_values['AIIC_recieve_corr'], 'tolist') else calc_values['AIIC_recieve_corr'],
                        'AIIC_Normalized_recieve': calc_values['AIIC_Normalized_recieve'].tolist() if hasattr(calc_values['AIIC_Normalized_recieve'], 'tolist') else calc_values['AIIC_Normalized_recieve'],
                        'AIIC_contour_val': calc_values['AIIC_contour_val'],
                        'AIIC_contour_result': calc_values['AIIC_contour_result'].tolist() if hasattr(calc_values['AIIC_contour_result'], 'tolist') else calc_values['AIIC_contour_result'],
                        'room_vol': calc_values['room_vol'],
                        'sabines': calc_values['sabines'].tolist() if hasattr(calc_values['sabines'], 'tolist') else calc_values['sabines']
                    }
                    logger.debug(f"Created AIIC entry: {aiic_entry}")
                    aiic_data.append(aiic_entry)
                except KeyError as e:
                    logger.error(f"Missing key in AIIC data: {e}")
                    logger.error(f"Available keys: {list(calc_values.keys())}")
                
                # Generate AIIC plot
                try:
                    IIC_curve = [2,2,2,2,2,2,1,0,-1,-2,-3,-6,-9,-12,-15,-18]
                    IIC_contour_final = [val + (110 - calc_values['AIIC_contour_val']) for val in IIC_curve]
                    freq_series = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 
                                 1000, 1250, 1600, 2000, 2500, 3150]
                    
                    plot_path = plot_curves(
                        frequencies=freq_series,
                        y_label='Sound Pressure Level (dB)',
                        ref_curve=np.array(IIC_contour_final),
                        field_curve=np.array(calc_values['AIIC_Normalized_recieve']),
                        ref_label=f'AIIC {calc_values["AIIC_contour_val"]} Contour',
                        field_label='Absorption Normalized Impact Sound Pressure Level, AIIC (dB)'
                    )
                    
                    # Move plot to output directory
                    plot_name = f"{test_label}_AIIC_plot.png"
                    final_plot_path = output_dir / plot_name
                    Path(plot_path).rename(final_plot_path)
                    logger.info(f"Saved plot to: {final_plot_path}")
                except Exception as e:
                    logger.error(f"Error generating AIIC plot: {e}")
                
            elif test_type == TestType.ASTC:
                logger.info(f"Processing ASTC data for {test_label}")
                try:
                    astc_entry = {
                        'Test_Label': test_label,
                        'NR_val': calc_values['NR_val'].tolist() if hasattr(calc_values['NR_val'], 'tolist') else calc_values['NR_val'],
                        'ATL_val': calc_values['ATL_val'].tolist() if hasattr(calc_values['ATL_val'], 'tolist') else calc_values['ATL_val'],
                        'ASTC_final_val': calc_values['ASTC_final_val'],
                        'ASTC_contour_val': calc_values['ASTC_contour_val'].tolist() if hasattr(calc_values['ASTC_contour_val'], 'tolist') else calc_values['ASTC_contour_val'],
                        'ASTC_recieve_corr': calc_values['ASTC_recieve_corr'].tolist() if hasattr(calc_values['ASTC_recieve_corr'], 'tolist') else calc_values['ASTC_recieve_corr'],
                        'sabines': calc_values['sabines'].tolist() if hasattr(calc_values['sabines'], 'tolist') else calc_values['sabines'],
                        'room_vol': calc_values['room_vol']
                    }
                    logger.debug(f"Created ASTC entry: {astc_entry}")
                    astc_data.append(astc_entry)
                except KeyError as e:
                    logger.error(f"Missing key in ASTC data: {e}")
                    logger.error(f"Available keys: {list(calc_values.keys())}")
                
            elif test_type == TestType.NIC:
                logger.info(f"Processing NIC data for {test_label}")
                try:
                    nic_entry = {
                        'Test_Label': test_label,
                        'NR_val': calc_values['NR_val'].tolist() if hasattr(calc_values['NR_val'], 'tolist') else calc_values['NR_val'],
                        'NIC_contour_val': calc_values['NIC_contour_val'],
                        'NIC_final_val': calc_values['NIC_final_val'],
                        'NIC_recieve_corr': calc_values['NIC_recieve_corr'].tolist() if hasattr(calc_values['NIC_recieve_corr'], 'tolist') else calc_values['NIC_recieve_corr'],
                        'sabines': calc_values['sabines'].tolist() if hasattr(calc_values['sabines'], 'tolist') else calc_values['sabines'],
                        'room_vol': calc_values['room_vol']
                    }
                    logger.debug(f"Created NIC entry: {nic_entry}")
                    nic_data.append(nic_entry)
                except KeyError as e:
                    logger.error(f"Missing key in NIC data: {e}")
                    logger.error(f"Available keys: {list(calc_values.keys())}")
    
    # Save DataFrames to CSV
    try:
        if aiic_data:
            logger.debug(f"AIIC data to save: {aiic_data}")
            aiic_df = pd.DataFrame(aiic_data)
            aiic_path = output_dir / 'aiic_results.csv'
            aiic_df.to_csv(aiic_path, index=False)
            logger.info(f"Saved AIIC results to: {aiic_path}")
            logger.debug(f"AIIC DataFrame shape: {aiic_df.shape}")
            logger.debug(f"AIIC DataFrame columns: {aiic_df.columns}")
            logger.debug(f"AIIC DataFrame head:\n{aiic_df.head()}")
            
        if astc_data:
            logger.debug(f"ASTC data to save: {astc_data}")
            astc_df = pd.DataFrame(astc_data)
            astc_path = output_dir / 'astc_results.csv'
            astc_df.to_csv(astc_path, index=False)
            logger.info(f"Saved ASTC results to: {astc_path}")
            logger.debug(f"ASTC DataFrame shape: {astc_df.shape}")
            logger.debug(f"ASTC DataFrame columns: {astc_df.columns}")
            logger.debug(f"ASTC DataFrame head:\n{astc_df.head()}")
            
        if nic_data:
            logger.debug(f"NIC data to save: {nic_data}")
            nic_df = pd.DataFrame(nic_data)
            nic_path = output_dir / 'nic_results.csv'
            nic_df.to_csv(nic_path, index=False)
            logger.info(f"Saved NIC results to: {nic_path}")
            logger.debug(f"NIC DataFrame shape: {nic_df.shape}")
            logger.debug(f"NIC DataFrame columns: {nic_df.columns}")
            logger.debug(f"NIC DataFrame head:\n{nic_df.head()}")
            
    except Exception as e:
        logger.error(f"Error saving CSV files: {e}")
        logger.error(f"Error type: {type(e)}")
        import traceback
        logger.error(f"Traceback:\n{traceback.format_exc()}")
        raise 