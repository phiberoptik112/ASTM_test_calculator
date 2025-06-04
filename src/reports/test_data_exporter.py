import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from pathlib import Path
from typing import Dict, List, Union
from src.core.data_processor import TestType

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
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Create DataFrames for each test type
    aiic_data = []
    astc_data = []
    nic_data = []
    
    # Process each test
    for test_label, test_types in test_data_collection.items():
        for test_type, test_info in test_types.items():
            test_data = test_info['test_data']
            calc_values = test_data.calculated_values
            
            if test_type == TestType.AIIC:
                aiic_data.append({
                    'Test_Label': test_label,
                    'NR_val': calc_values['NR_val'],
                    'AIIC_recieve_corr': calc_values['AIIC_recieve_corr'],
                    'AIIC_Normalized_recieve': calc_values['AIIC_Normalized_recieve'],
                    'AIIC_contour_val': calc_values['AIIC_contour_val'],
                    'AIIC_contour_result': calc_values['AIIC_contour_result'],
                    'room_vol': calc_values['room_vol'],
                    'sabines': calc_values['sabines']
                })
                
                # Generate AIIC plot
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
                Path(plot_path).rename(output_dir / plot_name)
                
            elif test_type == TestType.ASTC:
                astc_data.append({
                    'Test_Label': test_label,
                    'NR_val': calc_values['NR_val'],
                    'ATL_val': calc_values['ATL_val'],
                    'ASTC_final_val': calc_values['ASTC_final_val'],
                    'ASTC_contour_val': calc_values['ASTC_contour_val'],
                    'ASTC_recieve_corr': calc_values['ASTC_recieve_corr'],
                    'sabines': calc_values['sabines'],
                    'room_vol': calc_values['room_vol']
                })
                
            elif test_type == TestType.NIC:
                nic_data.append({
                    'Test_Label': test_label,
                    'NR_val': calc_values['NR_val'],
                    'NIC_contour_val': calc_values['NIC_contour_val'],
                    'NIC_final_val': calc_values['NIC_final_val'],
                    'NIC_recieve_corr': calc_values['NIC_recieve_corr'],
                    'sabines': calc_values['sabines'],
                    'room_vol': calc_values['room_vol']
                })
    
    # Save DataFrames to CSV
    if aiic_data:
        pd.DataFrame(aiic_data).to_csv(output_dir / 'aiic_results.csv', index=False)
    if astc_data:
        pd.DataFrame(astc_data).to_csv(output_dir / 'astc_results.csv', index=False)
    if nic_data:
        pd.DataFrame(nic_data).to_csv(output_dir / 'nic_results.csv', index=False) 