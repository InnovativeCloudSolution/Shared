import sys
import random
import os
import time
import requests
from cw_rpa import Logger, Input, HttpClient, ResultLevel

sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

log = Logger()
http_client = HttpClient()
input = Input()
log.info("Imports completed successfully")

cwpsa_base_url = "https://aus.myconnectwise.net"
cwpsa_base_url_path = "/v4_6_release/apis/3.0"

data_to_log = {}
bot_name = "CWPSA - Agreement custom field to invoice"
log.info("Static variables set")

def record_result(log, level, message):
    log.result_message(level, f"[{bot_name}]: {message}")
    if level == ResultLevel.WARNING:
        data_to_log["status_result"] = "Fail"
    elif level == ResultLevel.SUCCESS:
        if "status_result" not in data_to_log or data_to_log["status_result"] != "Fail":
            data_to_log["status_result"] = "Success"

def execute_api_call(log, http_client, method, endpoint, data=None, retries=5, integration_name=None, headers=None, params=None):
    base_delay = 5
    log.info(f"Executing API call: {method.upper()} {endpoint}")
    for attempt in range(retries):
        try:
            if integration_name:
                response = (
                    getattr(http_client.third_party_integration(integration_name), method)(url=endpoint, json=data)
                    if data else getattr(http_client.third_party_integration(integration_name), method)(url=endpoint)
                )
            else:
                request_args = {"url": endpoint}
                if params:
                    request_args["params"] = params
                if headers:
                    request_args["headers"] = headers
                if data:
                    if (headers and headers.get("Content-Type") == "application/x-www-form-urlencoded"):
                        request_args["data"] = data
                    else:
                        request_args["json"] = data
                response = getattr(requests, method)(**request_args)

            if 200 <= response.status_code < 300:
                return response
            elif response.status_code in [429, 503]:
                retry_after = response.headers.get("Retry-After")
                wait_time = int(retry_after) if retry_after else base_delay * (2 ** attempt) + random.uniform(0, 3)
                log.warning(f"Rate limit exceeded. Retrying in {wait_time:.2f} seconds")
                time.sleep(wait_time)
            elif 400 <= response.status_code < 500:
                if response.status_code == 404:
                    log.warning(f"Skipping non-existent resource [{endpoint}]")
                    return None
                log.error(f"Client error Status: {response.status_code}, Response: {response.text}")
                return response
            elif 500 <= response.status_code < 600:
                log.warning(f"Server error Status: {response.status_code}, attempt {attempt + 1} of {retries}")
                time.sleep(base_delay * (2 ** attempt) + random.uniform(0, 3))
            else:
                log.error(f"Unexpected response Status: {response.status_code}, Response: {response.text}")
                return response

        except Exception as e:
            log.exception(e, f"Exception during API call to {endpoint}")
            return None
    return None

def get_invoice(log, http_client, cwpsa_base_url, cwpsa_base_url_path, invoice_id):
    log.info(f"Retrieving invoice [{invoice_id}]")
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/finance/invoices/{invoice_id}"
    response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
    if response and 200 <= response.status_code < 300:
        return response.json()
    return None

def get_agreement_additions(log, http_client, cwpsa_base_url, cwpsa_base_url_path, agreement_id):
    log.info(f"Retrieving additions for agreement [{agreement_id}]")
    additions = []
    page = 1
    page_size = 250
    while True:
        endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/finance/agreements/{agreement_id}/additions?pageSize={page_size}&page={page}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if not response or response.status_code != 200:
            break
        batch = response.json()
        if not batch:
            break
        additions.extend(batch)
        if len(batch) < page_size:
            break
        page += 1
    log.info(f"Retrieved [{len(additions)}] additions for agreement [{agreement_id}]")
    return additions

def get_invoice_products(log, http_client, cwpsa_base_url, cwpsa_base_url_path, invoice_id):
    log.info(f"Retrieving products for invoice [{invoice_id}]")
    products = []
    page = 1
    page_size = 250
    while True:
        endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/procurement/products?conditions=invoice/id={invoice_id}&pageSize={page_size}&page={page}"
        response = execute_api_call(log, http_client, "get", endpoint, integration_name="cw_psa")
        if not response or response.status_code != 200:
            break
        batch = response.json()
        if not batch:
            break
        products.extend(batch)
        if len(batch) < page_size:
            break
        page += 1
    log.info(f"Retrieved [{len(products)}] products for invoice [{invoice_id}]")
    return products

def find_custom_field_by_caption(custom_fields, caption):
    for field in custom_fields:
        if field.get("caption", "").lower() == caption.lower():
            return field
    return None

def update_product_custom_fields(log, http_client, cwpsa_base_url, cwpsa_base_url_path, product_id, custom_field_updates):
    log.info(f"Updating [{len(custom_field_updates)}] custom field(s) on product [{product_id}]")
    endpoint = f"{cwpsa_base_url}{cwpsa_base_url_path}/procurement/products/{product_id}"
    patch_data = [
        {
            "op": "replace",
            "path": "customFields",
            "value": custom_field_updates
        }
    ]
    response = execute_api_call(log, http_client, "patch", endpoint, data=patch_data, integration_name="cw_psa")
    if response and 200 <= response.status_code < 300:
        log.info(f"Successfully updated custom fields on product [{product_id}]")
        return True
    log.error(f"Failed to update custom fields on product [{product_id}]")
    return False

def main():
    try:
        try:
            invoice_id = input.get_value("InvoiceID_1773090910444")
            custom_field_1 = input.get_value("CustomField1_1773090911903")
            custom_field_2 = input.get_value("CustomField2_1773090914457")
            custom_field_3 = input.get_value("CustomField3_1773090913204")
            custom_field_4 = input.get_value("CustomField4_1773090916079")
        except Exception:
            record_result(log, ResultLevel.WARNING, "Failed to fetch input values")
            return

        invoice_id = invoice_id.strip() if invoice_id else ""
        custom_field_1 = custom_field_1.strip() if custom_field_1 else ""
        custom_field_2 = custom_field_2.strip() if custom_field_2 else ""
        custom_field_3 = custom_field_3.strip() if custom_field_3 else ""
        custom_field_4 = custom_field_4.strip() if custom_field_4 else ""

        log.info(f"Invoice ID = [{invoice_id}]")
        log.info(f"Custom Field 1 = [{custom_field_1}]")
        log.info(f"Custom Field 2 = [{custom_field_2}]")
        log.info(f"Custom Field 3 = [{custom_field_3}]")
        log.info(f"Custom Field 4 = [{custom_field_4}]")

        if not invoice_id:
            record_result(log, ResultLevel.WARNING, "Invoice ID is required but missing")
            return
        if not custom_field_1:
            record_result(log, ResultLevel.WARNING, "At least one custom field name is required")
            return

        field_names_to_copy = [f for f in [custom_field_1, custom_field_2, custom_field_3, custom_field_4] if f]
        log.info(f"Custom fields to copy: {field_names_to_copy}")

        invoice = get_invoice(log, http_client, cwpsa_base_url, cwpsa_base_url_path, invoice_id)
        if not invoice:
            record_result(log, ResultLevel.WARNING, f"Invoice [{invoice_id}] not found")
            return

        invoice_products = get_invoice_products(log, http_client, cwpsa_base_url, cwpsa_base_url_path, invoice_id)
        if not invoice_products:
            record_result(log, ResultLevel.WARNING, f"No products found on invoice [{invoice_id}]")
            return

        products_by_agreement = {}
        for product in invoice_products:
            agr_id = product.get("agreement", {}).get("id")
            if not agr_id:
                log.warning(f"Invoice product [{product.get('id')}] has no agreement reference - skipping")
                continue
            if agr_id not in products_by_agreement:
                products_by_agreement[agr_id] = []
            products_by_agreement[agr_id].append(product)

        log.info(f"Invoice products grouped into [{len(products_by_agreement)}] agreement(s): {list(products_by_agreement.keys())}")

        updated_count = 0
        skipped_count = 0
        agreements_processed = []

        for agreement_id, agreement_products in products_by_agreement.items():
            agreement_products.sort(key=lambda p: p.get("id", 0))
            log.info(f"Processing agreement [{agreement_id}] with [{len(agreement_products)}] invoice product(s)")

            additions = get_agreement_additions(log, http_client, cwpsa_base_url, cwpsa_base_url_path, agreement_id)
            if not additions:
                log.warning(f"No additions found on agreement [{agreement_id}] - skipping [{len(agreement_products)}] invoice product(s)")
                skipped_count += len(agreement_products)
                continue

            billable_additions = [a for a in additions if a.get("billCustomer") != "DoNotBill" and not a.get("cancelledDate")]
            billable_additions.sort(key=lambda a: a.get("sequenceNumber", 0))
            log.info(f"Agreement [{agreement_id}]: [{len(billable_additions)}] billable additions, [{len(agreement_products)}] invoice products")
            agreements_processed.append(agreement_id)

            if len(billable_additions) != len(agreement_products):
                log.warning(f"Agreement [{agreement_id}]: Count mismatch - [{len(billable_additions)}] additions vs [{len(agreement_products)}] products")

            match_count = min(len(billable_additions), len(agreement_products))
            for i in range(match_count):
                addition = billable_additions[i]
                product = agreement_products[i]

                addition_product_id = addition.get("product", {}).get("id")
                invoice_catalog_id = product.get("catalogItem", {}).get("id")
                addition_seq = addition.get("sequenceNumber", "?")
                product_seq = product.get("sequenceNumber", "?")
                product_id = product.get("id")

                if addition_product_id != invoice_catalog_id:
                    log.warning(f"Agreement [{agreement_id}] position [{i}]: product mismatch - addition product ID [{addition_product_id}] (seq {addition_seq}) != invoice catalog ID [{invoice_catalog_id}] (seq {product_seq}) - skipping")
                    skipped_count += 1
                    continue

                addition_custom_fields = addition.get("customFields", [])
                product_custom_fields = product.get("customFields", [])

                if not addition_custom_fields:
                    log.info(f"Agreement [{agreement_id}] position [{i}]: No custom fields on addition (seq {addition_seq}) - skipping")
                    continue

                fields_to_update = []
                for field_name in field_names_to_copy:
                    source_field = find_custom_field_by_caption(addition_custom_fields, field_name)
                    if not source_field:
                        log.info(f"Agreement [{agreement_id}] position [{i}]: Custom field [{field_name}] not found on addition (seq {addition_seq})")
                        continue

                    source_value = source_field.get("value")
                    target_field = find_custom_field_by_caption(product_custom_fields, field_name)
                    if not target_field:
                        log.warning(f"Agreement [{agreement_id}] position [{i}]: Custom field [{field_name}] not found on invoice product [{product_id}] (seq {product_seq})")
                        continue

                    target_field_id = target_field.get("id")
                    log.info(f"Agreement [{agreement_id}] position [{i}]: Copying [{field_name}] = [{source_value}] from addition (seq {addition_seq}) to product [{product_id}] (seq {product_seq})")
                    fields_to_update.append({"id": target_field_id, "value": source_value})

                if fields_to_update:
                    if update_product_custom_fields(log, http_client, cwpsa_base_url, cwpsa_base_url_path, product_id, fields_to_update):
                        updated_count += 1
                    else:
                        skipped_count += 1
                else:
                    log.info(f"Agreement [{agreement_id}] position [{i}]: No matching custom fields to update on product [{product_id}]")

        data_to_log["invoice_id"] = int(invoice_id)
        data_to_log["agreements_processed"] = agreements_processed
        data_to_log["total_invoice_products"] = len(invoice_products)
        data_to_log["updated"] = updated_count
        data_to_log["skipped"] = skipped_count
        data_to_log["fields_copied"] = field_names_to_copy

        if skipped_count > 0 and updated_count > 0:
            record_result(log, ResultLevel.SUCCESS, f"Updated [{updated_count}] product(s) on invoice [{invoice_id}], skipped [{skipped_count}] due to mismatch or errors")
        elif updated_count > 0:
            record_result(log, ResultLevel.SUCCESS, f"Successfully updated custom fields on [{updated_count}] product(s) for invoice [{invoice_id}]")
        elif skipped_count > 0:
            record_result(log, ResultLevel.WARNING, f"No products updated on invoice [{invoice_id}] - [{skipped_count}] skipped due to mismatch or errors")
        else:
            record_result(log, ResultLevel.SUCCESS, f"No custom fields needed updating on invoice [{invoice_id}]")

    except Exception as e:
        log.error(f"Unhandled error in main: {str(e)}")
        record_result(log, ResultLevel.WARNING, "Unhandled exception occurred during execution")
    finally:
        log.result_data(data_to_log)

if __name__ == "__main__":
    main()
