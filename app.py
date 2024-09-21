import os
import streamlit as st
import re

# Carregar variáveis de ambiente
client = OpenAI(api_key=os.getenv('OPENAI_API_KEY'))

from pydantic import BaseModel, Field
from openai import OpenAI
import pandas as pd

#Stremlit
import io
from io import StringIO
import xlsxwriter





def extraction():
  global content_filtered
  output_file_path = '/content/extracted_sections.txt'

  # Step 2: Find all sections labeled "REPORT- SV-A System Design Parameters"
  section_title = "REPORT- SV-A System Design Parameters"
  start_indices = [m.start() for m in re.finditer(section_title, content)]

  # Sections to exclude
  exclude_titles = [
      "REPORT- SS-C", "REPORT- SS-D", "REPORT- SS-H",
      "REPORT- SS-R", "REPORT- SS-K", "REPORT- SS-L", "REPORT- SS-I"
  ]

  # Step 3: Extract sections while excluding unwanted ones
  sections = []
  for i, start_index in enumerate(start_indices):
      # Find the next occurrence or the end of the content
      end_index = start_indices[i + 1] if i + 1 < len(start_indices) else len(content)
      section = content[start_index:end_index]

      # Check if the section contains any of the excluded titles
      if any(exclude_title in section for exclude_title in exclude_titles):
          continue  # Skip sections with excluded titles

      # Include only sections with "REPORT- SV-A System Design Parameters"
      if section_title in section:
          sections.append(section)

  # Step 4: Write the filtered sections to a new file
  with open(output_file_path, 'w', encoding='ISO-8859-1') as output_file:
      for section in sections:
          output_file.write(section)
          output_file.write('\n\n')  # Add a separator between sections

  print(f"Filtered sections have been saved to {output_file_path}")

  with open(output_file_path, 'r', encoding='ISO-8859-1') as file:
      content_filtered = file.read()



#Gen AI
client = OpenAI()
class System_eQUEST(BaseModel):
    system_name_s: str = Field(..., description="System name")
    system_type: str = Field(..., description="System type")
    altitude_factor: float = Field(..., description="Altitude factor")
    floor_area_sqft: float = Field(..., description="Floor area in square feet")
    max_people: int = Field(..., description="Maximum number of people")
    outside_air_ratio: float = Field(..., description="Outside air ratio")
    cooling_capacity_kbtu_hr: float = Field(..., description="Cooling capacity in KBTU/hr")
    sensible_heat_ratio: float = Field(..., description="Sensible heat ratio (SHR)")
    heating_capacity_kbtu_hr: float = Field(..., description="Heating capacity in KBTU/hr")
    cooling_eir_btu_btu: float = Field(..., description="Cooling energy efficiency ratio in BTU/BTU")
    heating_eir_btu_btu: float = Field(..., description="Heating energy efficiency ratio in BTU/BTU")
    heat_pump_supp_heat_kbtu_hr: float = Field(..., description="Heat pump supplementary heat in KBTU/hr")


class FanSystem(BaseModel):
    system_name_f: str = Field(..., description="System name")
    fan_type: str = Field(..., description="Fan type (e.g., SUPPLY, RETURN)")
    capacity_cfm: float = Field(..., description="Capacity in cubic feet per minute (CFM)")
    diversity_factor_frac: float = Field(..., description="Diversity factor (fraction)")
    power_demand_kw: float = Field(..., description="Power demand in kilowatts (kW)")
    fan_delta_t_f: float = Field(..., description="Fan delta temperature in Fahrenheit (°F)")
    static_pressure_in_water: float = Field(..., description="Static pressure in inches of water (in-H2O)")
    total_eff_frac: float = Field(..., description="Total efficiency (fraction)")
    mech_eff_frac: float = Field(..., description="Mechanical efficiency (fraction)")
    fan_placement: str = Field(..., description="Fan placement (e.g., BLOW-THRU, DRAW-THRU)")
    fan_control: str = Field(..., description="Fan control method (e.g., BY USER, AUTOMATIC)")
    max_fan_ratio_frac: float = Field(..., description="Maximum fan ratio (fraction)")
    min_fan_ratio_frac: float = Field(..., description="Minimum fan ratio (fraction)")


from typing import List, Union

#class SystemCollection(BaseModel):
#    systems: List[System_eQUEST] = Field(..., description="List of systems")

class SystemCollection(BaseModel):
    systems: List[Union[System_eQUEST, FanSystem]] = Field(..., description="List of systems (System_eQUEST and FanSystem)")

#class SystemCollection(BaseModel):
#    systems: List[Union[System_eQUEST, FanSystem, ZoneSystem]] = Field(..., description="List of systems (System_eQUEST, FanSystem, and ZoneSystem)")


def structured():
  global completion
  completion = client.beta.chat.completions.parse(
      model="gpt-4o-2024-08-06",
      temperature = 0,
      messages=[
          {"role": "system", "content": "You are an expert at structured data extraction. You will be given unstructured text from eQuest and should convert it into the given structure for all systems in the file."},
          {"role": "user", "content": f"{content_filtered}"}
      ],
      response_format=SystemCollection,
  )


#excel_file_path = "/content/df_system_equest.xlsx"
#df_system_equest.to_excel(excel_file_path, index=False)

#excel_file_path = "/content/df_fan_system.xlsx"
#df_fan_system.to_excel(excel_file_path, index=False)

# Create a BytesIO buffer to write the Excel file in memory.
buffer = io.BytesIO()


# Streamlit UI
st.title("eQuest data extractor")
st.write("Powered by GPT-4o")


uploaded_file = st.file_uploader("Load your .SIM file")
if uploaded_file is not None:
    # Convert the uploaded file content to a string, assuming it uses ISO-8859-1 encoding
    stringio = StringIO(uploaded_file.getvalue().decode('ISO-8859-1'))

    # Step 1: Read the file content
    content = stringio.read()

    extraction()
    structured()
    system_response = completion.choices[0].message.parsed
    response = system_response.systems

    system_equest_list = [item for item in response if isinstance(item, System_eQUEST)]
    fan_system_list = [item for item in response if isinstance(item, FanSystem)]
    # Convert to DataFrames
    df_system_equest = pd.DataFrame([item.dict() for item in system_equest_list])
    df_fan_system = pd.DataFrame([item.dict() for item in fan_system_list])


    # Create a Pandas Excel writer using XlsxWriter as the engine.
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
      # Write each dataframe to a different worksheet.
      df_system_equest.to_excel(writer, sheet_name='SystemEquest')
      df_fan_system.to_excel(writer, sheet_name='FanSystem')

    # Provide a download button for the user to download the file.
    st.download_button(
        label="Download Excel File",
        data=buffer.getvalue(),
        file_name="system_fan_data.xlsx",
        mime="application/vnd.ms-excel"
    )
