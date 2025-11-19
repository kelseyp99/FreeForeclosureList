
import React from "react";
import "../App.css";
import { Link } from "react-router-dom";


export default function OrangeCounty() {
  const [selectedCounty, setSelectedCounty] = React.useState("");
  const [selectedSaleType, setSelectedSaleType] = React.useState("");

  // Build report file name
  let reportSrc = "";
  if (selectedCounty && selectedSaleType) {
    reportSrc = `/reports/sales_report_${selectedCounty.toLowerCase().replace(/\s/g, "_")}_${selectedSaleType.toLowerCase().replace(/\s/g, "")}.html`;
  }

  return (
    <div className="container" style={{ display: 'flex', minHeight: '100vh' }}>
      <FloridaCountiesSidebar
        onSelectReport={(county, saleType) => {
          setSelectedCounty(county);
          setSelectedSaleType(saleType === 'UiPath' ? 'foreclosure' : 'taxdeed');
        }}
      />
      <div style={{ flex: 1, display: 'flex', flexDirection: 'column' }}>
        <main className="main-content" style={{ padding: '40px 32px 0 32px', flex: 1 }}>
          <Header />
          <h2>Florida County Foreclosure & Tax Deed Reports</h2>
          <div style={{ display: 'flex', flexDirection: 'row', gap: 40, alignItems: 'flex-start', marginBottom: 24 }}>
            <div style={{ flex: 1, minWidth: 320 }}>
              <strong>Hello. We're FreeForeclosureList.net</strong>
              <p>Welcome to FreeForeclosureList.net, your premier destination for accessing comprehensive real estate distressed property listings. Powered by cutting-edge AI and Robotic Process Automation, we revolutionize the way you explore foreclosure properties. Unlike traditional county foreclosure lists, we go above and beyond by curating additional insights sourced from the web, providing you with a one-stop solution for all your real estate investment needs.</p>
              <p>Understanding the demands of modern investors, we offer invaluable features such as direct links to various real estate platforms, county property appraisers, and clerks of court. Our platform delivers more than just basic information; we provide estimated property values, judgment amounts for foreclosure cases, and opening bid amounts for Tax Deed sales. This empowers you to gauge potential equity and focus your efforts efficiently. By identifying properties where lenders are likely to halt bidding at the judgment amount, we save you valuable time. Moreover, you may discover opportunities to connect with property owners who owe less than the judgment amount, opening avenues for direct purchase.</p>
              <p>In addition to our comprehensive foreclosure data, we also offer exclusive access to sales information from counties, including proprietary and hard-to-obtain lists.</p>
              <p><em>Please note that FreeForeclosureList.net is currently in its prototype stage. Expect significant enhancements and updates in the coming months and weeks as we strive to provide you with an unparalleled user experience.</em></p>
              {reportSrc && (
                <div style={{ marginTop: 32 }}>
                  <h3>{selectedCounty} County {selectedSaleType === 'foreclosure' ? 'Foreclosure' : 'Tax Deed'} Report</h3>
                  <iframe
                    src={reportSrc}
                    title="County Sales Report"
                    style={{ width: '100%', minHeight: 600, border: '1px solid #ccc', borderRadius: 8 }}
                  />
                </div>
              )}
            </div>
            <div style={{ flex: 1, minWidth: 320 }}>
              <h3>Upcoming Foreclosure Sales</h3>
              <div className="highlight-box">
                <p>Get the latest list of foreclosure and tax deed properties in Florida counties. Updated weekly!</p>
              </div>
            </div>
          </div>
        </main>
        <footer className="footer">
          <div>FreeForeclosureList.net</div>
          <div>©{new Date().getFullYear()} Florida County Foreclosure & Tax Deed List. All rights reserved.</div>
        </footer>
      </div>
    </div>
  );
}
