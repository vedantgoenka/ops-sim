#!/usr/bin/env python3
"""
Data analysis module for the ops-sim project.
This module provides functionality for analyzing and updating the master Excel file.
"""

# Standard library imports
from pathlib import Path
from typing import Optional, Dict, Any, Union
import re

# Third-party imports
import pandas as pd

# Local imports
import config


class DataAnalyzer:
    """Class for analyzing and updating the master Excel file.
    
    This class provides methods for:
    - Loading data from the master Excel file
    - Adding current price information based on history
    - Adding capacity allocation information based on history
    - Adding initial and final batch size information based on history
    """
    
    def __init__(self, master_file_path: Optional[Union[str, Path]] = None) -> None:
        """Initialize the DataAnalyzer.
        
        Args:
            master_file_path: Path to the master Excel file. If None, uses the path from config.
        """
        if master_file_path is None:
            self.master_file_path = config.MASTER_FILE
        else:
            self.master_file_path = Path(master_file_path)
        
        self.data: Optional[Dict[str, pd.DataFrame]] = None
        self.load_data()
    
    def load_data(self) -> None:
        """Load all sheets from the master Excel file.
        
        Raises:
            FileNotFoundError: If the master file doesn't exist
            ValueError: If the file is not a valid Excel file
        """
        try:
            if not self.master_file_path.exists():
                raise FileNotFoundError(f"Master file not found: {self.master_file_path}")
            
            self.data = pd.read_excel(
                self.master_file_path,
                sheet_name=None,
                keep_default_na=False,
                na_filter=False
            )
            print("Successfully loaded data from master file")
        except Exception as e:
            print(f"Error loading data: {str(e)}")
            self.data = None
            raise
    
    def get_sheet(self, sheet_name: str) -> Optional[pd.DataFrame]:
        """Get a specific sheet from the loaded data.
        
        Args:
            sheet_name: Name of the sheet to retrieve.
            
        Returns:
            The requested sheet as a DataFrame, or None if not found.
            
        Raises:
            ValueError: If no data is loaded.
        """
        if self.data is None:
            raise ValueError("No data loaded. Call load_data() first.")
        return self.data.get(sheet_name)
    
    def _extract_price_updates(self, history: pd.DataFrame) -> pd.DataFrame:
        """Extract price updates from the History sheet.
        
        Args:
            history: DataFrame containing the History sheet data.
            
        Returns:
            DataFrame containing price updates with Day and Price columns.
        """
        price_updates = history[history['Description'].str.contains('price', case=False)].copy()
        if not price_updates.empty:
            price_updates['Price'] = price_updates['Description'].str.extract(r'\$(\d+)').astype(int)
            price_updates = price_updates[['Day', 'Price']].sort_values('Day', ascending=False)
            price_updates = price_updates.drop_duplicates('Day')
            price_updates = price_updates.sort_values('Day')
        return price_updates
    
    def _extract_capacity_updates(self, history: pd.DataFrame) -> pd.DataFrame:
        """Extract capacity allocation updates from the History sheet.
        
        Args:
            history: DataFrame containing the History sheet data.
            
        Returns:
            DataFrame containing capacity updates with Day and Allocation columns.
        """
        capacity_updates = history[history['Description'].str.contains('capacity allocation', case=False)].copy()
        if not capacity_updates.empty:
            capacity_updates['Allocation'] = capacity_updates['Description'].str.extract(r'to (\d+\.?\d*)').astype(float)
            capacity_updates['Allocation'] = capacity_updates['Allocation'].round(2)
            capacity_updates = capacity_updates[['Day', 'Allocation']].sort_values('Day', ascending=False)
            capacity_updates = capacity_updates.drop_duplicates('Day')
            capacity_updates = capacity_updates.sort_values('Day')
        return capacity_updates
    
    def _extract_initial_batch_size_updates(self, history: pd.DataFrame) -> pd.DataFrame:
        """Extract initial batch size updates from the History sheet.
        
        Args:
            history: DataFrame containing the History sheet data.
            
        Returns:
            DataFrame containing initial batch size updates with Day and InitialBatchSize columns.
        """
        # Look for entries like "Updated initial standard batch size to 125 units."
        batch_updates = history[history['Description'].str.contains('initial standard batch size', case=False)].copy()
        if not batch_updates.empty:
            batch_updates['InitialBatchSize'] = batch_updates['Description'].str.extract(r'to (\d+) units').astype(int)
            batch_updates = batch_updates[['Day', 'InitialBatchSize']].sort_values('Day', ascending=False)
            batch_updates = batch_updates.drop_duplicates('Day')
            batch_updates = batch_updates.sort_values('Day')
        return batch_updates
    
    def _extract_final_batch_size_updates(self, history: pd.DataFrame) -> pd.DataFrame:
        """Extract final batch size updates from the History sheet.
        
        Args:
            history: DataFrame containing the History sheet data.
            
        Returns:
            DataFrame containing final batch size updates with Day and FinalBatchSize columns.
        """
        # Look for entries like "Updated final standard batch size to 28 units."
        batch_updates = history[history['Description'].str.contains('final standard batch size', case=False)].copy()
        if not batch_updates.empty:
            batch_updates['FinalBatchSize'] = batch_updates['Description'].str.extract(r'to (\d+) units').astype(int)
            batch_updates = batch_updates[['Day', 'FinalBatchSize']].sort_values('Day', ascending=False)
            batch_updates = batch_updates.drop_duplicates('Day')
            batch_updates = batch_updates.sort_values('Day')
        return batch_updates
    
    def add_current_price(self) -> None:
        """Add a column with current product price to the Standard sheet based on History updates.
        
        This method:
        1. Extracts price updates from the History sheet
        2. Updates the Standard sheet with the most recent price for each day
        3. Saves the updated data back to the master file
        
        Raises:
            ValueError: If required sheets are not found
            IOError: If there's an error saving the updated data
        """
        history = self.get_sheet('History')
        standard = self.get_sheet('Standard')
        
        if history is None or standard is None:
            raise ValueError("Required sheets not found")
        
        # Extract price updates
        price_updates = self._extract_price_updates(history)
        
        # Initialize with default price
        standard['Current Price'] = int(config.DEFAULT_PRICE)
        
        if not price_updates.empty:
            min_update_day = price_updates['Day'].min()
            
            # Update price for each day based on the most recent update
            for day in sorted(standard['Day'].unique()):
                if day >= min_update_day:
                    applicable_updates = price_updates[price_updates['Day'] <= day]
                    if not applicable_updates.empty:
                        latest_price = int(applicable_updates.iloc[-1]['Price'])
                        standard.loc[standard['Day'] == day, 'Current Price'] = latest_price
        
        # Save the updated data
        try:
            with pd.ExcelWriter(self.master_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                standard.to_excel(writer, sheet_name='Standard', index=False)
            print("Successfully added Current Price column to Standard sheet")
        except Exception as e:
            raise IOError(f"Error saving updated data: {str(e)}")
    
    def add_capacity_allocation(self) -> None:
        """Add a column with current capacity allocation percentage to the Standard sheet.
        
        This method:
        1. Extracts capacity allocation updates from the History sheet
        2. Updates the Standard sheet with the most recent allocation for each day
        3. Saves the updated data back to the master file
        
        Raises:
            ValueError: If required sheets are not found
            IOError: If there's an error saving the updated data
        """
        history = self.get_sheet('History')
        standard = self.get_sheet('Standard')
        
        if history is None or standard is None:
            raise ValueError("Required sheets not found")
        
        # Initialize with default allocation
        standard['Capacity Allocation %'] = round(float(config.DEFAULT_ALLOCATION), 2)
        
        # Extract capacity updates
        capacity_updates = self._extract_capacity_updates(history)
        
        if not capacity_updates.empty:
            min_update_day = capacity_updates['Day'].min()
            
            # Update allocation for each day based on the most recent update
            for day in sorted(standard['Day'].unique()):
                if day >= min_update_day:
                    applicable_updates = capacity_updates[capacity_updates['Day'] <= day]
                    if not applicable_updates.empty:
                        latest_allocation = round(float(applicable_updates.iloc[-1]['Allocation']), 2)
                        standard.loc[standard['Day'] == day, 'Capacity Allocation %'] = latest_allocation
        
        # Save the updated data
        try:
            with pd.ExcelWriter(self.master_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                standard.to_excel(writer, sheet_name='Standard', index=False)
            print("Successfully added Capacity Allocation % column to Standard sheet")
        except Exception as e:
            raise IOError(f"Error saving updated data: {str(e)}")
    
    def add_batch_sizes(self) -> None:
        """Add columns with initial and final batch sizes to the Standard sheet.
        
        This method:
        1. Extracts batch size updates from the History sheet
        2. Updates the Standard sheet with the most recent batch sizes for each day
        3. Saves the updated data back to the master file
        
        Raises:
            ValueError: If required sheets are not found
            IOError: If there's an error saving the updated data
        """
        history = self.get_sheet('History')
        standard = self.get_sheet('Standard')
        
        if history is None or standard is None:
            raise ValueError("Required sheets not found")
        
        # Initialize with default batch sizes
        standard['Initial Batch Size'] = int(config.DEFAULT_INITIAL_BATCH_SIZE)
        standard['Final Batch Size'] = int(config.DEFAULT_FINAL_BATCH_SIZE)
        
        # Extract batch size updates
        initial_batch_updates = self._extract_initial_batch_size_updates(history)
        final_batch_updates = self._extract_final_batch_size_updates(history)
        
        # Update initial batch sizes
        if not initial_batch_updates.empty:
            min_update_day = initial_batch_updates['Day'].min()
            
            for day in sorted(standard['Day'].unique()):
                if day >= min_update_day:
                    applicable_updates = initial_batch_updates[initial_batch_updates['Day'] <= day]
                    if not applicable_updates.empty:
                        latest_batch_size = int(applicable_updates.iloc[-1]['InitialBatchSize'])
                        standard.loc[standard['Day'] == day, 'Initial Batch Size'] = latest_batch_size
        
        # Update final batch sizes
        if not final_batch_updates.empty:
            min_update_day = final_batch_updates['Day'].min()
            
            for day in sorted(standard['Day'].unique()):
                if day >= min_update_day:
                    applicable_updates = final_batch_updates[final_batch_updates['Day'] <= day]
                    if not applicable_updates.empty:
                        latest_batch_size = int(applicable_updates.iloc[-1]['FinalBatchSize'])
                        standard.loc[standard['Day'] == day, 'Final Batch Size'] = latest_batch_size
        
        # Save the updated data
        try:
            with pd.ExcelWriter(self.master_file_path, engine='openpyxl', mode='a', if_sheet_exists='replace') as writer:
                standard.to_excel(writer, sheet_name='Standard', index=False)
            print("Successfully added Initial and Final Batch Size columns to Standard sheet")
        except Exception as e:
            raise IOError(f"Error saving updated data: {str(e)}")


def main() -> None:
    """Main entry point for the script."""
    try:
        analyzer = DataAnalyzer()
        analyzer.add_current_price()
        analyzer.add_capacity_allocation()
        analyzer.add_batch_sizes()
    except Exception as e:
        print(f"Error in analysis: {str(e)}")
        raise


if __name__ == "__main__":
    main() 