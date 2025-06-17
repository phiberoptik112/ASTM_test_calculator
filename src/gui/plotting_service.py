"""
Plotting Service for ASTM Test Calculator

This module handles all plotting and visualization logic for test data.
Extracted from main_window.py to separate plotting logic from UI logic.
"""

import traceback
import matplotlib.pyplot as plt
import numpy as np
import tempfile
from typing import Dict, List, Optional, Any, Tuple
from src.core.data_processor import TestType, TestData, AIICTestData, ASTCTestData, NICTestData


class PlottingService:
    """Service class for handling all plotting operations"""
    
    def __init__(self, debug_mode: bool = False):
        self.debug_mode = debug_mode
        # Standard frequency bands for ASTM testing
        self.standard_freqs = [100, 125, 160, 200, 250, 315, 400, 500, 630, 800, 1000, 1250, 1600, 2000, 2500, 3150, 4000]
    
    def create_test_plot(self, test_type: TestType, test_obj: TestData, freq_data: Dict[str, Any], 
                        calculated_values: Dict[str, Any]) -> Optional[str]:
        """
        Create a plot for the specified test type and return the temporary file path
        
        Args:
            test_type: The type of test (AIIC, ASTC, NIC)
            test_obj: The test data object
            freq_data: Processed frequency data
            calculated_values: Calculated test values
            
        Returns:
            Path to temporary plot file or None if failed
        """
        try:
            # Create figure with subplots
            fig, (ax1, ax2) = plt.subplots(2, 1, figsize=(12, 10))
            
            success = False
            if test_type == TestType.AIIC:
                success = self._create_aiic_plot(ax1, ax2, test_obj, freq_data, calculated_values)
            elif test_type == TestType.ASTC:
                success = self._create_astc_plot(ax1, ax2, test_obj, freq_data, calculated_values)
            elif test_type == TestType.NIC:
                success = self._create_nic_plot(ax1, ax2, test_obj, freq_data, calculated_values)
            
            if not success:
                plt.close(fig)
                return None
            
            # Configure common plot settings
            self._configure_plot_axes(ax1, ax2)
            
            # Save plot to temporary file
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                plt.savefig(temp_file.name, dpi=300, bbox_inches='tight')
                plt.close(fig)
                return temp_file.name
                
        except Exception as e:
            print(f"Error creating {test_type} plot: {str(e)}")
            traceback.print_exc()
            plt.close('all')  # Clean up any open figures
            return None
    
    def _create_aiic_plot(self, ax1: plt.Axes, ax2: plt.Axes, test_obj: AIICTestData, 
                         freq_data: Dict[str, Any], results: Dict[str, Any]) -> bool:
        """Create AIIC-specific plots"""
        try:
            if self.debug_mode:
                print("Creating AIIC analysis plots")
            
            # Plot raw data on ax1
            ax1.set_title('Raw Test Data - AIIC')
            ax1.set_xlabel('Frequency (Hz)')
            ax1.set_ylabel('Sound Pressure Level (dB)')
            
            # Plot source data
            ax1.plot(self.standard_freqs, freq_data['source'], 
                    label='Source Room', color='green', marker='s')
            
            # Plot position data
            for i, pos_data in enumerate(results['positions'], 1):
                ax1.plot(self.standard_freqs, pos_data,
                        linestyle='--', label=f'AIIC - Tapping Position {i}')
            
            # Plot background data
            ax1.plot(self.standard_freqs, freq_data['background'],
                    label='Background', color='gray', marker='x')
            
            # Plot analysis on ax2
            ax2.set_title('AIIC Analysis')
            ax2.set_xlabel('Frequency (Hz)')
            ax2.set_ylabel('Sound Pressure Level (dB)')
            
            # Plot normalized receive levels
            ax2.plot(self.standard_freqs, results['AIIC_Normalized_recieve'], 
                    label='Normalized Impact Sound Level', 
                    color='blue', marker='o')
            
            # Calculate and plot AIIC contour
            IIC_curve = [2, 2, 2, 2, 2, 2, 1, 0, -1, -2, -3, -6, -9, -12, -15, -18]
            IIC_contour_final = [val + (110 - results['AIIC_contour_val']) for val in IIC_curve]
            
            ax2.plot(self.standard_freqs, IIC_contour_final, 
                    label=f'AIIC {results["AIIC_contour_val"]} Contour', 
                    color='red', linestyle='--')
            
            # Plot background check exceptions
            for i, (freq, level, exception) in enumerate(zip(self.standard_freqs, 
                                                           results['AIIC_Normalized_recieve'],
                                                           results['AIIC_exceptions_backcheck'])):
                if exception == '1':
                    ax2.axvline(x=freq, color='orange', alpha=0.6, linestyle=':', 
                               label='Background Check Exception' if i == 0 else "")
            
            # Add results text box
            textstr = f'AIIC: {results["AIIC_contour_val"]}'
            props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
            ax2.text(0.05, 0.95, textstr, transform=ax2.transAxes, 
                    verticalalignment='top', bbox=props)
            
            if self.debug_mode:
                print("Successfully created AIIC analysis plots")
            return True
            
        except Exception as e:
            print(f"Error creating AIIC analysis plots: {str(e)}")
            traceback.print_exc()
            return False
    
    def _create_astc_plot(self, ax1: plt.Axes, ax2: plt.Axes, test_obj: ASTCTestData, 
                         freq_data: Dict[str, Any], results: Dict[str, Any]) -> bool:
        """Create ASTC-specific plots"""
        try:
            if self.debug_mode:
                print("Creating ASTC analysis plots")
            
            # Plot raw data on ax1
            ax1.set_title('Raw Test Data - ASTC')
            ax1.set_xlabel('Frequency (Hz)')
            ax1.set_ylabel('Sound Pressure Level (dB)')
            
            # Plot source and receive data
            ax1.plot(self.standard_freqs, freq_data['source'], 
                    label='Source Room', color='green', marker='s')
            ax1.plot(self.standard_freqs, freq_data['receive'], 
                    label='Receive Room', color='blue', marker='o')
            ax1.plot(self.standard_freqs, freq_data['background'],
                    label='Background', color='gray', marker='x')
            
            # Plot analysis on ax2
            ax2.set_title('ASTC Analysis')
            ax2.set_xlabel('Frequency (Hz)')
            ax2.set_ylabel('Sound Level (dB)')
            
            # Plot ATL values
            ax2.plot(self.standard_freqs, results['ATL_val'], 
                    label='Apparent Transmission Loss (ATL)', 
                    color='blue', marker='o')
            
            # Plot ASTC contour
            ax2.plot(self.standard_freqs, results['ASTC_contour_val'], 
                    label=f'ASTC {results["ASTC_final_val"]} Contour', 
                    color='red', linestyle='--')
            
            # Add results text box
            textstr = (f'ASTC: {results["ASTC_final_val"]}\n'
                      f'Room Volume: {results["room_vol"]:.1f} m³')
            props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
            ax2.text(0.05, 0.95, textstr, transform=ax2.transAxes, 
                    verticalalignment='top', bbox=props)
            
            if self.debug_mode:
                print("Successfully created ASTC analysis plots")
            return True
            
        except Exception as e:
            print(f"Error creating ASTC analysis plots: {str(e)}")
            traceback.print_exc()
            return False
    
    def _create_nic_plot(self, ax1: plt.Axes, ax2: plt.Axes, test_obj: NICTestData, 
                        freq_data: Dict[str, Any], results: Dict[str, Any]) -> bool:
        """Create NIC-specific plots"""
        try:
            if self.debug_mode:
                print("Creating NIC analysis plots")
            
            # Plot raw data on ax1
            ax1.set_title('Raw Test Data - NIC')
            ax1.set_xlabel('Frequency (Hz)')
            ax1.set_ylabel('Sound Pressure Level (dB)')
            
            # Plot source and receive data
            ax1.plot(self.standard_freqs, freq_data['source'], 
                    label='Source Room', color='green', marker='s')
            ax1.plot(self.standard_freqs, freq_data['receive'], 
                    label='Receive Room', color='blue', marker='o')
            ax1.plot(self.standard_freqs, freq_data['background'],
                    label='Background', color='gray', marker='x')
            
            # Plot analysis on ax2
            ax2.set_title('NIC Analysis')
            ax2.set_xlabel('Frequency (Hz)')
            ax2.set_ylabel('Sound Level (dB)')
            
            # Plot NR values
            ax2.plot(self.standard_freqs, results['NR_val'], 
                    label=f"Noise Reduction (NIC = {results['NIC_contour_val']})", 
                    color='blue', marker='o')
            
            # Plot NIC reference curve
            ref_curve = np.array([-16, -13, -10, -7, -4, -1, 0, 1, 2, 3, 4, 4, 4, 4, 4, 4, 4]) + results['NIC_contour_val']
            ax2.plot(self.standard_freqs, ref_curve, 
                    label='NIC Reference Curve', 
                    color='red', linestyle='--')
            
            # Add results text box
            textstr = (f"NIC: {results['NIC_contour_val']}\n"
                      f"Sabines: {np.mean(results['sabines']):.1f}\n"
                      f"Room Volume: {results['room_vol']:.1f} m³")
            props = dict(boxstyle='round', facecolor='wheat', alpha=0.5)
            ax2.text(0.05, 0.95, textstr, transform=ax2.transAxes, 
                    verticalalignment='top', bbox=props)
            
            if self.debug_mode:
                print("Successfully created NIC analysis plots")
            return True
            
        except Exception as e:
            print(f"Error creating NIC analysis plots: {str(e)}")
            traceback.print_exc()
            return False
    
    def _configure_plot_axes(self, ax1: plt.Axes, ax2: plt.Axes) -> None:
        """Configure common plot settings for both axes"""
        try:
            # Set log scale and configure x-axis for both plots
            for ax in [ax1, ax2]:
                ax.grid(True)
                ax.set_xscale('log')
                ax.set_xticks(self.standard_freqs)
                ax.set_xticklabels([str(f) for f in self.standard_freqs])
                ax.legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize='small')
            
            # Adjust layout to prevent legend cutoff
            plt.tight_layout()
            
        except Exception as e:
            print(f"Error configuring plot axes: {str(e)}")
            traceback.print_exc()
    
    def create_simple_plot(self, x_data: List[float], y_data: List[float], 
                          title: str, xlabel: str, ylabel: str, label: str = None) -> Optional[str]:
        """
        Create a simple line plot and return temporary file path
        
        Args:
            x_data: X-axis data
            y_data: Y-axis data
            title: Plot title
            xlabel: X-axis label
            ylabel: Y-axis label
            label: Line label for legend
            
        Returns:
            Path to temporary plot file or None if failed
        """
        try:
            fig, ax = plt.subplots(figsize=(10, 6))
            
            ax.plot(x_data, y_data, marker='o', label=label if label else '')
            ax.set_title(title)
            ax.set_xlabel(xlabel)
            ax.set_ylabel(ylabel)
            ax.grid(True)
            
            if label:
                ax.legend()
            
            # Save plot to temporary file
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                plt.savefig(temp_file.name, dpi=300, bbox_inches='tight')
                plt.close(fig)
                return temp_file.name
                
        except Exception as e:
            print(f"Error creating simple plot: {str(e)}")
            traceback.print_exc()
            plt.close('all')
            return None
    
    def cleanup_temp_plots(self, plot_paths: List[str]) -> None:
        """
        Clean up temporary plot files
        
        Args:
            plot_paths: List of temporary file paths to clean up
        """
        import os
        
        for path in plot_paths:
            try:
                if os.path.exists(path):
                    os.remove(path)
                    if self.debug_mode:
                        print(f"Cleaned up temporary plot: {path}")
            except Exception as e:
                print(f"Error cleaning up plot file {path}: {str(e)}")
    
    def get_standard_frequencies(self) -> List[int]:
        """Return the standard 1/3 octave frequency bands used in ASTM testing"""
        return self.standard_freqs.copy()