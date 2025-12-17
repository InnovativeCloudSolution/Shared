import requests
import json
import sys
from datetime import datetime

def lookup_cve(cve_id):
    url = f"https://nvd.nist.gov/vuln/detail/{cve_id}"
    api_url = f"https://services.nvd.nist.gov/rest/json/cves/2.0?cveId={cve_id}"
    
    try:
        response = requests.get(api_url, timeout=10)
        if response.status_code == 200:
            data = response.json()
            
            if 'vulnerabilities' in data and len(data['vulnerabilities']) > 0:
                vuln = data['vulnerabilities'][0]['cve']
                
                cvss_score = 0
                severity = "Unknown"
                
                if 'metrics' in vuln:
                    if 'cvssMetricV31' in vuln['metrics']:
                        cvss_data = vuln['metrics']['cvssMetricV31'][0]['cvssData']
                        cvss_score = cvss_data.get('baseScore', 0)
                        severity = cvss_data.get('baseSeverity', 'Unknown')
                    elif 'cvssMetricV2' in vuln['metrics']:
                        cvss_data = vuln['metrics']['cvssMetricV2'][0]['cvssData']
                        cvss_score = cvss_data.get('baseScore', 0)
                        severity = "HIGH" if cvss_score >= 7.0 else "MEDIUM" if cvss_score >= 4.0 else "LOW"
                
                description = ""
                if 'descriptions' in vuln:
                    description = vuln['descriptions'][0]['value']
                
                published = vuln.get('published', '')
                
                return {
                    "cve_id": cve_id,
                    "severity": severity,
                    "cvss_score": cvss_score,
                    "description": description[:500],
                    "published_date": published,
                    "url": url
                }
        
        return {"cve_id": cve_id, "severity": "Unknown", "cvss_score": 0, "description": "CVE data not found", "url": url}
    
    except Exception as e:
        return {"cve_id": cve_id, "severity": "Error", "cvss_score": 0, "description": str(e), "url": url}

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(json.dumps({"error": "CVE ID required"}))
        sys.exit(1)
    
    cve_id = sys.argv[1]
    result = lookup_cve(cve_id)
    print(json.dumps(result))

