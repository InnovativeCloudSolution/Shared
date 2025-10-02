# Synopsis
# --------
# Bot Name       : Update Opportunity Custom Property
# Description    : This bot calculates and updates various custom properties for a given
#                  opportunity ID. It handles:
#                  - Total Contract Value
#                  - Total One-Time Value
#                  - Total Contract Margin
#                  - First Year Revenue
#                  - First Year Margin
#                  It logs individual product values and updates these metrics in the
#                  Opportunity Custom fields dynamically. Handles creation of new fields
#                  if they do not exist.
# Developed by   : Prakash Chalgeri
# Created Date   : 12/12/2024 (dd/mm/yyyy)
# Updated On     : 09/01/2025 (dd/mm/yyyy)
# Version        : 2.2
# Jira           : RPA-1312

# Import System libraries required for Pre-Check
import subprocess
import os
import io
import json
import sys
import base64
import hashlib
import time
import urllib
import random
import traceback
import random
import string
from cw_rpa import Logger, Input, HttpClient
from datetime import datetime

# Declare global variables & initialize logger and HTTP client
log = Logger()
http_client = HttpClient()
input = Input()
CYCLES_MAPPING = {
    "monthly": 12,
    "quarterly": 4,
    "annually": 1
}
# Function block started
def generate_error_reference():
    error_ref = random.randint(10000, 99999)
    return f'#{error_ref}'

def log_stdout(msg):
    error_code = generate_error_reference()
    sys.stdout.write(f"\nError reference: {error_code}\nError: {msg}\nTraceback : {traceback.format_exc()}".replace('\n', '\\n'))
    log.error(f"An internal error occured, error reference: {error_code}")
    log.result_failed_message(f"An internal error occured, error reference: {error_code}")

# Function to calculate total contract value
def get_total_contract_value(base_url, opp_id):
    try:
        # Construct the URL
        product_endpoint = f"{base_url}/procurement/products?conditions=opportunity/id={opp_id}&pagesize=1000"
       
        # Make the API call
        response = http_client.third_party_integration("cw_psa").get(url=product_endpoint)
        api_call_data = response.json()

        # Check for a successful response
        if response.status_code != 200:
            log.error(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            log.result_failed_message(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            return

        if not api_call_data:
            log.result_failed_message(f"No products found for opportunity ID: {opp_id}.")
            log.error(f"No products found for opportunity ID: {opp_id}.")
            return

        log.info(f"Successfully fetched products for opportunity ID: {opp_id}.")
       
        # Initialize variables
        total_contract_value = 0
        contract_values = {}

        # Process each product
        for product in api_call_data:
            try:
                # Check for recurringRevenue property and filter
                recurring_revenue = product.get("recurring", {}).get("recurringRevenue", 0)
                if recurring_revenue > 0:
                    product_name = product.get("catalogItem", {}).get("identifier", "Unknown_Product")
                    cycles = product.get("recurring", {}).get("cycles", 1)
                    quantity = product.get("quantity", 1)

                    # Calculate the contract value
                    contract_value = recurring_revenue * cycles * quantity
                    contract_values[f"contract_value_{product_name}"] = contract_value

                    # Update total contract value
                    total_contract_value += contract_value

                    log.info(f"Calculated contract value for product '{product_name}': {contract_value}")

            except Exception as e:
                log.error(f"Failed to process product: {product}. Error: {str(e)}")
                log.result_failed_message(f"Failed to process product: {product}. Error: {str(e)}")

        # Add total contract value to the dictionary
        contract_values["total_contract_value"] = total_contract_value
        log.info(f"Total contract value calculated: {total_contract_value}")

        return contract_values

    except Exception as e:
        log.stdout(e, f"An error occurred while fetching products for opportunity ID: {opp_id}.")        
#Check if a custom property exists in opportunities.
def check_custom_property_exist(base_url,custom_property_name):
    
    try:
        # Construct the API endpoint URL
        custom_property_endpoint = f"{base_url}/system/userDefinedFields?conditions=screenId='my_opportunities'&pagesize=1000"

        # Make the API call to fetch custom properties
        response = http_client.third_party_integration("cw_psa").get(url=custom_property_endpoint)
        

        # Handle non-200 responses
        if response.status_code != 200:
            log.error(f"Failed to fetch custom properties. Status code: {response.status_code}, Response: {response.content}")
            return {"isPropertyPresent": False, "sequenceNumbers": {}}

        # Check if the response content is empty
        if not response.content.strip():
            log.info("No custom properties found in opportunities.")
            return {"isPropertyPresent": False, "sequenceNumbers": {}}

        # Parse the API response
        api_call_data = response.json()

        # If response is blank or empty JSON
        if not api_call_data:
            log.info("No custom properties found in opportunities.")
            return {"isPropertyPresent": False, "sequenceNumbers": {}}
        
        log.info("Successfully fetched custom properties for opportunities.")

        # Variables to track results
        is_property_present = False
        sequence_numbers_by_pod = {}

        # Iterate over the response and check for 'custom_property_name'
        for custom_property in api_call_data:
            sequence_number = custom_property.get("sequenceNumber", None)
            pod_id = custom_property.get("podId", None)
            caption = custom_property.get("caption", "")

            # Check if the 'caption' matches 'custom_property_name'
            if caption.lower() == custom_property_name.lower():
                is_property_present = True

            # Group sequence numbers by podId
            if sequence_number is not None and pod_id is not None:
                if pod_id not in sequence_numbers_by_pod:
                    sequence_numbers_by_pod[pod_id] = []
                sequence_numbers_by_pod[pod_id].append(sequence_number)

        # Return the results
        result = {
            "isPropertyPresent": is_property_present,
            "sequenceNumbers": sequence_numbers_by_pod
        }

        log.info(f"Custom property '{custom_property_name}' found: {is_property_present}")
        #log.info(f"Sequence numbers grouped by podId: {sequence_numbers_by_pod}")
        return result

    except Exception as e:
        log.stdout(e, "An error occurred while checking for custom properties.")
        return {"isPropertyPresent": False, "sequenceNumbers": {}}
#Update the custom property 'custom_property_name' for a specific opportunity.
def update_custom_property_value(base_url, opportunity_id, total_contract_value,custom_property_name):
    
    try:
        # First API call: Get the opportunity details to fetch the custom property index
        opportunity_endpoint = f"{base_url}/sales/opportunities/{opportunity_id}"
        response = http_client.third_party_integration("cw_psa").get(url=opportunity_endpoint)
        
        if response.status_code != 200:
            log.error(f"Failed to fetch opportunity details. Status code: {response.status_code}, Response: {response.content}")
            return False
        
        opportunity_data = response.json()

        custom_fields = opportunity_data.get("customFields", [])
        
        # Find the index of 'custom_property_name' custom property
        index_number = None
        for index, field in enumerate(custom_fields):
            if field.get("caption") == custom_property_name:
                index_number = index
                break
        
        if index_number is None:
            log.error(f"Custom property '{custom_property_name}' not found under opportunity.")
            return False
        
        # Second API call: Update the custom property value using a PATCH request
        patch_endpoint = f"{base_url}/sales/opportunities/{opportunity_id}"
        patch_body = [{
            "op": "replace",
            "path": f"customFields/{index_number}/value",
            "value": str(total_contract_value)
        }]
        
        patch_response = http_client.third_party_integration("cw_psa").patch(
            url=patch_endpoint,
            json=patch_body
        )
        
        if patch_response.status_code == 200:
            log.info(f"Successfully updated '{custom_property_name}' to {total_contract_value} for opportunity ID {opportunity_id}.")
            log.result_success_message(f"Successfully updated '{custom_property_name}' to {total_contract_value} for opportunity ID {opportunity_id}.")
            return True
        else:
            log.error(f"Failed to update '{custom_property_name}'. Status code: {patch_response.status_code}, Response: {patch_response.content}")
            log.result_failed_message(f"Failed to update '{custom_property_name}'. Status code: {patch_response.status_code}, Response: {patch_response.content}")
            return False
    except Exception as e:
        log.stdout(e, "An error occurred while updating the custom property value.")
        return False
#Create a new custom property if it does not already exist.
def create_custom_property(base_url, custom_property_name, sequence_numbers):
    try:
        # Define the podId to be used (hardcoded for now)
        pod_id = get_pod_id(base_url)

        # Determine the sequenceNumber for the podId
        if not sequence_numbers:
            # Case 1: sequenceNumbers is empty
            sequence_number = 1
        elif pod_id not in sequence_numbers:
            # Case 2: sequenceNumbers has data but podId is not in it
            sequence_number = 1
        else:
            # Case 3: sequenceNumbers has data and podId is present
            existing_numbers = sorted(sequence_numbers[pod_id])
            sequence_number = existing_numbers[-1] + 1 if existing_numbers else 1
        location_info = get_location_and_department_info(base_url)
        # Create the request body
        create_body = {
            "caption": custom_property_name,
            "displayOnScreenFlag": True,
            "entryTypeIdentifier": "EntryField",
            "fieldTypeIdentifier": "Text",
            "sequenceNumber": sequence_number,
            "helpText": custom_property_name,
            "podId": pod_id,
            "readOnlyFlag": False,
            "screenId": "my_opportunities",
            "locationIds": location_info['locationIds'],
            "businessUnitIds": location_info['businessUnitIds']
        }

        # API endpoint for creating a custom property
        create_endpoint = f"{base_url}/system/userDefinedFields"

        # Make the API call
        response = http_client.third_party_integration("cw_psa").post(
            url=create_endpoint,
            json=create_body
        )

        if response.status_code == 201:
            log.info(f"Successfully created custom property '{custom_property_name}' with sequence number {sequence_number}.")
            log.result_success_message(f"Successfully created custom property '{custom_property_name}' with sequence number {sequence_number}.")
            return True
        else:
            log.error(f"Failed to create custom property. Status code: {response.status_code}, Response: {response.content}")
            log.result_failed_message(f"Failed to create custom property. Status code: {response.status_code}, Response: {response.content}")
            return False

    except Exception as e:
        log.stdout(e, "An error occurred while creating the custom property.")
        return False
#Fetch the location IDs and business unit IDs from the API.
def get_location_and_department_info(base_url):
    try:
        # API endpoint for locations
        endpoint = f"{base_url}/system/locations"

        # Make the API call
        response = http_client.third_party_integration("cw_psa").get(url=endpoint)

        if response.status_code == 200:
            api_data = response.json()

            # Extract location IDs
            location_ids = [item['id'] for item in api_data if 'id' in item] or [0]

            # Extract business unit IDs (departmentIds)
            business_unit_ids = [
                dep_id
                for item in api_data if 'departmentIds' in item
                for dep_id in item['departmentIds']
            ] or [0]

            log.info("Successfully fetched location and department information.")
            return {
                "locationIds": location_ids,
                "businessUnitIds": business_unit_ids
            }

        else:
            log.error(f"Failed to fetch location and department info. Status code: {response.status_code}, Response: {response.content}")
            return {
                "locationIds": [0],
                "businessUnitIds": [0]
            }

    except Exception as e:
        log.stdout(e, "An error occurred while fetching location and department information.")
        return {
            "locationIds": [0],
            "businessUnitIds": [0]
        }
#Function to fetch the podId for 'opportunities_opportunity'.
def get_pod_id(base_url):
    try:
        # API endpoint for fetching pod information
        pod_endpoint = f"{base_url}/system/reports/pod"
        
        # Make the API call
        response = http_client.third_party_integration("cw_psa").get(url=pod_endpoint)
        
        # Check for successful response
        if response.status_code != 200:
            log.info(f"Failed to fetch pod information. Status code: {response.status_code}, Response: {response.content}")
            return 2  # Default podId if API call fails
        
        # Parse the API response
        pod_data = response.json()
        row_values = pod_data.get("row_values", [])
        
        # Iterate through row_values to find 'opportunities_opportunity'
        for row in row_values:
            if len(row) > 1 and row[1] == "opportunities_opportunity":
                log.info(f"Found podId for 'opportunities_opportunity': {row[0]}")
                return row[0]
        
        # If 'opportunities_opportunity' is not found, return default
        log.info("poId with name 'opportunities_opportunity' not found in pod data. Using default podId: 2")
        return 2

    except Exception as e:
        log.stdout(e, "An error occurred while fetching the podId.")
        return 2  # Default podId in case of an exception
# Function to calculate total one-time value
def get_total_one_time_value(base_url, opp_id):

    try:
        # Construct the URL
        product_endpoint = f"{base_url}/procurement/products?conditions=opportunity/id={opp_id}&pagesize=1000"
       
        # Make the API call
        response = http_client.third_party_integration("cw_psa").get(url=product_endpoint)
       
        # Check for API response status
        if response.status_code != 200:
            log.error(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            return {"total_one_time_value": 0}
       
        # Parse the API response
        api_call_data = response.json()
        if not api_call_data:
            log.info(f"No products found for opportunity ID: {opp_id}.")
            log.result_success_message(f"No products found for opportunity ID: {opp_id}.")
            return {"total_one_time_value": 0}
       
        log.info(f"Successfully fetched products for opportunity ID: {opp_id}.")
       
        # Initialize total one-time value
        total_one_time_value = 0
       
        # Iterate over products and calculate the total one-time value
        for product in api_call_data:
            try:
                ext_price = product.get("extPrice", 0)
                if ext_price > 0:
                    total_one_time_value += ext_price
            except Exception as e:
                log.error(f"Failed to process product: {product}. Error: {str(e)}")
       
        log.info(f"Total One-Time Value calculated: {total_one_time_value}")
        return {"total_one_time_value": total_one_time_value}
   
    except Exception as e:
        log.stdout(e, f"An error occurred while calculating the Total One-Time Value for opportunity ID: {opp_id}.")
        return {"total_one_time_value": 0}
#Calculate the total contract margin for a given opportunity.
def get_total_contract_margin(base_url, opp_id):
    try:
        # Construct the API endpoint URL
        product_endpoint = f"{base_url}/procurement/products?conditions=opportunity/id={opp_id}&pagesize=1000"

        # Make the API call
        response = http_client.third_party_integration("cw_psa").get(url=product_endpoint)
       
        # Parse the response
        api_call_data = response.json()

        # Check for a successful response
        if response.status_code != 200:
            log.error(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            log.result_failed_message(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            return      

        if not api_call_data:
            log.result_failed_message(f"No products found for opportunity ID: {opp_id}.")
            log.error(f"No products found for opportunity ID: {opp_id}.")
            return

        log.info(f"Successfully fetched products for opportunity ID: {opp_id}.")

        # Initialize variables
        total_contract_margin = 0
        contract_margins = {}

        # Process each product
        for product in api_call_data:
            try:
                # Check for Recurring Revenue property
                recurring_revenue = product.get("recurring", {}).get("recurringRevenue", 0)
                recurring_cost = product.get("recurring", {}).get("recurringCost", 0)

                if recurring_revenue > 0:
                    # Calculate margin per product
                    margin_per_unit = recurring_revenue - recurring_cost
                    cycles = product.get("recurring", {}).get("cycles", 1)
                    quantity = product.get("quantity", 1)

                    # Calculate the contract margin
                    product_margin = margin_per_unit * cycles * quantity

                    # Store the product margin
                    product_name = product.get("catalogItem", {}).get("identifier", "Unknown_Product")
                    contract_margins[f"margin_{product_name}"] = product_margin

                    # Update total contract margin
                    total_contract_margin += product_margin

                    log.info(f"Calculated contract margin for product '{product_name}': {product_margin}")
            except Exception as e:
                log.error(f"Failed to process product: {product}. Error: {str(e)}")
                log.result_failed_message(f"Failed to process product: {product}. Error: {str(e)}")

        # Add total contract margin to the dictionary
        contract_margins["total_contract_margin"] = total_contract_margin

        log.info(f"Total contract margin calculated: {total_contract_margin}")

        return contract_margins
    except Exception as e:
        log_stdout(f"An error occurred while fetching products for opportunity ID: {opp_id}. {e}")
#Calculate the First Year Revenue for a given opportunity ID.
def get_first_year_revenue(base_url, opp_id):
    try:
        # Construct the URL to fetch products
        product_endpoint = f"{base_url}/procurement/products?conditions=opportunity/id={opp_id}&pagesize=1000"

        # Make the API call
        response = http_client.third_party_integration("cw_psa").get(url=product_endpoint)
        if response.status_code != 200:
            log.error(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            log.result_failed_message(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            return {}

        api_call_data = response.json()
        if not api_call_data:
            log.info(f"No products found for opportunity ID: {opp_id}.")
            log.result_success_message(f"No products found for opportunity ID: {opp_id}.")
            return {}

        log.info(f"Successfully fetched products for opportunity ID: {opp_id}.")

        # Initialize variables
        total_first_year_revenue = 0
        first_year_revenues = {}

        # Process each product
        for product in api_call_data:
            try:
                # Extract relevant values
                recurring_revenue = product.get("recurring", {}).get("recurringRevenue", 0)
                bill_cycle_name = product.get("recurring", {}).get("billCycle", {}).get("name", "").lower()
                quantity = product.get("quantity", 1)
                cycles = CYCLES_MAPPING.get(bill_cycle_name,1)
                # Calculate revenue only for products with recurringRevenue > 0
                if recurring_revenue > 0:
                    product_name = product.get("catalogItem", {}).get("identifier", "Unknown_Product")
                    first_year_revenue = recurring_revenue * cycles * quantity

                    # Log individual product revenue
                    first_year_revenues[f"first_year_revenue_{product_name}"] = first_year_revenue
                    total_first_year_revenue += first_year_revenue

                    log.info(f"Calculated first year revenue for product '{product_name}': {first_year_revenue}")

            except Exception as e:
                log.error(f"Failed to process product: {product}. Error: {str(e)}")
                log.result_failed_message(f"Failed to process product: {product}. Error: {str(e)}")

        # Add total first year revenue to the dictionary
        first_year_revenues["total_first_year_revenue"] = total_first_year_revenue
        log.info(f"Total first year revenue calculated: {total_first_year_revenue}")

        return first_year_revenues

    except Exception as e:
        log.stdout(e, f"An error occurred while calculating first year revenue for opportunity ID: {opp_id}.")
        return {}
#Calculate the First Year Margin for the given opportunity.
def get_first_year_margin(base_url, opp_id):
    try:
        # Construct the URL
        product_endpoint = f"{base_url}/procurement/products?conditions=opportunity/id={opp_id}&pagesize=1000"

        # Make the API call
        response = http_client.third_party_integration("cw_psa").get(url=product_endpoint)
        api_call_data = response.json()

        # Check for a successful response
        if response.status_code != 200:
            log.error(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            log.result_failed_message(f"Failed to fetch products. Status code: {response.status_code}, Response Content: {response.content}")
            return {}

        if not api_call_data:
            log.result_failed_message(f"No products found for opportunity ID: {opp_id}.")
            log.error(f"No products found for opportunity ID: {opp_id}.")
            return {}

        log.info(f"Successfully fetched products for opportunity ID: {opp_id}.")

        # Initialize variables
        total_first_year_margin = 0
        product_margins = {}

        # Process each product
        for product in api_call_data:
            try:
                # Check for recurringRevenue property and filter
                recurring_revenue = product.get("recurring", {}).get("recurringRevenue", 0)
                recurring_cost = product.get("recurring", {}).get("recurringCost", 0)

                if recurring_revenue > 0:
                    product_name = product.get("catalogItem", {}).get("identifier", "Unknown_Product")
                    bill_cycle = product.get("recurring", {}).get("billCycle", {}).get("name", "").lower()
                    quantity = product.get("quantity", 1)
                    cycles = CYCLES_MAPPING.get(bill_cycle,1)
                    # Calculate the margin for the product
                    margin = (recurring_revenue - recurring_cost) * cycles * quantity
                    product_margins[f"margin_{product_name}"] = margin

                    # Update total margin
                    total_first_year_margin += margin
                    log.info(f"Calculated margin for product '{product_name}': {margin}")

            except Exception as e:
                log.error(f"Failed to process product: {product}. Error: {str(e)}")
                log.result_failed_message(f"Failed to process product: {product}. Error: {str(e)}")

        # Add total margin to the dictionary
        product_margins["total_first_year_margin"] = total_first_year_margin
        log.info(f"Total First Year Margin calculated: {total_first_year_margin}")

        return product_margins

    except Exception as e:
        log.stdout(e, f"An error occurred while calculating First Year Margin for opportunity ID: {opp_id}.")
        return {}
# Helper Function
def handle_custom_property(base_url, opportunity_id, fetch_function, property_name, value_key):
    try:
        log.info(f"Starting process for custom property: '{property_name}'")
        # Step 1: Fetch the property value
        result = fetch_function(base_url, opportunity_id)
        if not result or result.get(value_key, 0) == 0:
            log.info(f"No data found for '{property_name}', skipping property update.")
            log.result_success_message(f"No data found for '{property_name}', skipping property update.")
            return

        value_to_update = result[value_key]
        log.info(f"'{property_name}': {value_to_update}")

        # Step 2: Check if the custom property exists
        property_status = check_custom_property_exist(base_url, property_name)
        if not property_status:
            log.result_failed_message(f"Failed to retrieve custom property '{property_name}'.")
            return

        # Step 3: Update or create the custom property
        if property_status["isPropertyPresent"]:
            success = update_custom_property_value(base_url, opportunity_id, value_to_update, property_name)
            if not success:
                log.error(f"Failed to update custom property '{property_name}'.")
                log.result_failed_message(f"Failed to update custom property '{property_name}'.")
        else:
            property_created = create_custom_property(base_url, property_name, property_status["sequenceNumbers"])
            if property_created:
                update_custom_property_value(base_url, opportunity_id, value_to_update, property_name)
            else:
                log.error(f"Failed to create custom property '{property_name}'.")
                log.result_failed_message(f"Failed to create custom property '{property_name}'.")

    except Exception as e:
        log.stdout(f"An error occurred while processing '{property_name}': {e}")
    finally:
        log.info(f"Completed process for custom property: '{property_name}'")
# Function block end

def main():
    log.info("Bot execution has started.")
    try:
        # Fetch input values
        opportunity_id = input.get_value("opportunityId_1734101260163")
        contract_property_name = "Total Contract Value"
        one_time_property_name = "Total One-Time Value"
        contract_margin_property_name = "Total Contract Margin"
        revenue_property_name = "First Year Revenue"
        margin_property_name = "First Year Margin"

        # Validate input
        if not opportunity_id:
            log.result_failed_message("Opportunity ID is required.")
            exit()

    except Exception as e:
        log_stdout(f"An error occurred while getting input values: {e}")
        return

    try:
        # Base URL
        cwpsa_base_url = "https://staging.connectwisedev.com/v4_6_release/apis/3.0"

        # Process each custom property
        properties_to_process = [
            {
                "function": get_total_contract_value,
                "property_name": contract_property_name,
                "value_key": "total_contract_value",
            },
            {
                "function": get_total_one_time_value,
                "property_name": one_time_property_name,
                "value_key": "total_one_time_value",
            },
            {
                "function": get_total_contract_margin,
                "property_name": contract_margin_property_name,
                "value_key": "total_contract_margin",
            },
            {
                "function": get_first_year_revenue,
                "property_name": revenue_property_name,
                "value_key": "total_first_year_revenue",
            },
            {
                "function": get_first_year_margin,
                "property_name": margin_property_name,
                "value_key": "total_first_year_margin",
            },
        ]
        for prop in properties_to_process:
            handle_custom_property(
                base_url=cwpsa_base_url,
                opportunity_id=opportunity_id,
                fetch_function=prop["function"],
                property_name=prop["property_name"],
                value_key=prop["value_key"],
            )
    except Exception as e:
        log_stdout(f"An error occurred while executing the bot: {e}")

    log.info("Bot execution has ended.")

# Entry point
if __name__ == "__main__":
    main()