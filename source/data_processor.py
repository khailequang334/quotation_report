import os
import pandas as pd
from typing import Tuple, Dict, Any


class DataProcessor:
    
    def __init__(self, input_file_path: str, config: Dict[str, Any]):
        self.input_file_path = input_file_path
        self.config = config
        self.partner_name = self._extract_partner_name()
    
    def _extract_partner_name(self) -> str:
        filename = os.path.basename(self.input_file_path)
        return filename[:(filename.find('.')-1)]
    
    def _get_quotation_data(self, sheet: str, skip: int) -> Tuple[pd.DataFrame, pd.DataFrame]:
        df = pd.read_excel(self.input_file_path, sheet_name=sheet, skiprows=skip)
        data20ft = df.iloc[:, [0,2,4,6,7,8,9,10,11]]
        data40ft = df.iloc[:, [0,3,5,6,7,8,9,10,12]]
        return data20ft, data40ft

    def _calculate_total_cost(self, data: pd.DataFrame) -> pd.DataFrame:
        port = data.iloc[:,0].str.upper()
        cost = data.iloc[:,1:9]
        
        for i in range(len(cost.columns)):
            cost.isetitem(i, pd.to_numeric(cost.iloc[:,i], errors='coerce'))
        
        cost.dropna(subset=[cost.columns[0]], inplace=True)
        cost = cost[cost[cost.columns[0]] != 0]
        
        total_cost = cost.sum(axis=1, numeric_only=True)
        
        return pd.DataFrame().assign(PORT=port, TOTALCOST=total_cost)

    def _calculate_average_port_cost(self, data: pd.DataFrame) -> pd.DataFrame:
        data.dropna(subset=[data.columns[1]], inplace=True)
        avg_data = data.groupby(data.columns[0])[data.columns[1]].mean().reset_index()
        return avg_data

    def _prepare_forwarder_data(self, min_cost: pd.DataFrame, output_file: str, 
                              sheet: str, skip: int) -> pd.DataFrame:
        df = pd.read_excel(output_file, sheet_name=sheet, skiprows=skip)
        
        pod = df.loc[:, 'POD'].str.upper()
        
        cost_mapping = dict(zip(min_cost[min_cost.columns[0]], min_cost[min_cost.columns[1]]))
        
        partner_costs = pod.map(cost_mapping)
        
        fwd_data = pd.DataFrame({
            'POD': pod,
            'PARTNER': self.partner_name,
            'COST': partner_costs
        })
        
        fwd_data = fwd_data.dropna(subset=['COST'])
        
        return fwd_data

    def Run(self, output_file: str, sheet_20ft: str, sheet_40ft: str) -> Dict[str, Any]:
        input_sheet = self.config['quotation']['input_sheet']
        
        data20ft, data40ft = self._get_quotation_data(input_sheet, 0)

        sum20 = self._calculate_total_cost(data20ft)
        sum40 = self._calculate_total_cost(data40ft)

        min_sum20 = self._calculate_average_port_cost(sum20)
        min_sum40 = self._calculate_average_port_cost(sum40)

        fwd_data_20 = self._prepare_forwarder_data(
            min_sum20, output_file, sheet_20ft, 3
        )
        fwd_data_40 = self._prepare_forwarder_data(
            min_sum40, output_file, sheet_40ft, 3
        )

        return {
            '20ft': {
                'partner': self.partner_name,
                'data': fwd_data_20,
                'sheet': sheet_20ft
            },
            '40ft': {
                'partner': self.partner_name,
                'data': fwd_data_40,
                'sheet': sheet_40ft
            }
        } 