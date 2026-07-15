/***************************************************************** *
UptoSkills Onboarding Automation * File : MailTemplates.gs * Version : 5.2
*****************************************************************/

function getCommonFooter() {
  return `
    <hr style="border: none; border-top: 1px solid #e2e8f0; margin: 30px 0;" />
    <table style="width: 100%; border-collapse: collapse;">
      <tr>
        <td style="padding: 0; vertical-align: top;">
          <p style="font-size: 14px; color: #64748b; line-height: 1.6; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 0 0 12px 0;">
            If you have any queries or require further operational assistance, feel free to contact us.<br/><br/>
            <span style="font-size: 12px; text-transform: uppercase; letter-spacing: 1px; color: #94a3b8; display: block; margin-bottom: 4px;">Authorized Signatory</span>
            <strong style="color: #0f172a; font-size: 15px;">Onboarding Team | HR</strong><br/>
            <span style="color: #0a192f; font-weight: 600;">${CONFIG.ORGANIZATION_NAME}</span>
          </p>
          
          <p style="font-size: 13px; color: #64748b; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; margin: 12px 0 0 0;">
            <a href="mailto:${CONFIG.HR_EMAIL}" style="color: #0a192f; text-decoration: none; font-weight: 600;">${CONFIG.HR_EMAIL}</a>
            <span style="color: #cbd5e1; margin: 0 8px;">|</span>
            <span style="font-size: 12px;">${CONFIG.ADDRESS}</span>
          </p>
        </td>
      </tr>
    </table>
    
    <div style="margin-top: 25px; padding: 16px; background-color: #f8fafc; border-radius: 6px; border-left: 3px solid #0a192f; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif;">
      <p style="font-size: 13px; font-weight: 600; color: #0a192f; margin: 0 0 6px 0; letter-spacing: 0.3px;">Let's Make Freshers Employable!</p>
      <p style="font-size: 12px; color: #64748b; margin: 0; line-height: 1.5;">
        Corporate Portal: <a href="${CONFIG.WEBSITE}" style="color: #0a192f; text-decoration: none; font-weight: 500;">uptoskills.com</a><br/>
        Network Channels: 
        <a href="${CONFIG.LINKEDIN}" style="color: #64748b; text-decoration: underline;">LinkedIn</a> &bull; 
        <a href="${CONFIG.INSTAGRAM}" style="color: #64748b; text-decoration: underline;">Instagram</a> &bull; 
        <a href="${CONFIG.YOUTUBE}" style="color: #64748b; text-decoration: underline;">YouTube</a> &bull; 
        <a href="${CONFIG.WHATSAPP_CHANNEL}" style="color: #64748b; text-decoration: underline;">WhatsApp Channel</a>
      </p>
    </div>
  `; 
}

function validationFailedTemplate(name, reasons, editUrl) {
  return `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 30px; border: 1px solid #e2e8f0; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
      <div style="background: linear-gradient(135deg, #0a192f 0%, #1e293b 100%); color: #ffffff; padding: 24px; text-align: center; border-radius: 8px; font-size: 18px; font-weight: 600; letter-spacing: 0.5px;">
        Onboarding Data Verification Action Required
      </div>
      <p style="margin-top: 25px; font-size: 15px;">Dear <strong>${escapeHtml(name)}</strong>,</p>
      <p style="font-size: 15px; color: #475569;">Your onboarding profile data has been processed by our verification systems. A systematic mismatch requires your immediate corrective attention:</p>
      
      <div style="background-color: #fef2f2; border: 1px solid #fee2e2; border-left: 4px solid #ef4444; padding: 18px; margin: 20px 0; border-radius: 6px;">
        <strong style="color: #991b1b; font-size: 14px; text-transform: uppercase; letter-spacing: 0.5px; display: block; margin-bottom: 6px;">Identified Discrepancies:</strong>
        <pre style="font-family: 'SFMono-Regular', Consolas, 'Liberation Mono', Menlo, monospace; white-space: pre-line; margin: 0; color: #7f1d1d; font-size: 13px; line-height: 1.5;">${escapeHtml(reasons)}</pre>
      </div>
      
      <p style="margin: 30px 0; text-align: center;">
        <a href="${editUrl}" target="_blank"
           style="background-color: #0a192f; color: #ffffff;
           padding: 14px 28px; text-decoration: none;
           border-radius: 6px; display: inline-block; font-weight: 600; font-size: 14px; letter-spacing: 0.3px; transition: all 0.2s ease;">
           Modify Onboarding Records
        </a>
      </p>
      ${getCommonFooter()}
    </div>
  `; 
}

/** * High-End Modern Intern Welcome View
 */ 
function internWelcomeTemplate(data, groupInfo) {
  return `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 30px; border: 1px solid #e2e8f0; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
      <div style="background: linear-gradient(135deg, #0a192f 0%, #16a34a 100%); color: #ffffff; padding: 24px; text-align: center; border-radius: 8px; font-size: 20px; font-weight: 600; letter-spacing: 0.5px;">
        Internship Allocation Matrix Verified
      </div>
      <p style="margin-top: 25px; font-size: 15px;">Dear <strong>${escapeHtml(data.name)}</strong>,</p>
      <p style="font-size: 15px; color: #475569;">Your company credential profile has been successfully generated. Your functional cohort deployment details are finalized below:</p>
      
      <table style="width: 100%; border-collapse: separate; border-spacing: 0; margin: 25px 0; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
        <tr>
          <td style="padding: 14px 18px; background-color: #f8fafc; width: 40%; font-weight: 600; color: #475569; border-bottom: 1px solid #e2e8f0; font-size: 14px;">Intern Code Identity</td>
          <td style="padding: 14px 18px; font-weight: 700; color: #0a192f; border-bottom: 1px solid #e2e8f0; font-size: 14px; letter-spacing: 0.5px;">${escapeHtml(data.code)}</td>
        </tr>
        <tr>
          <td style="padding: 14px 18px; background-color: #f8fafc; font-weight: 600; color: #475569; border-bottom: 1px solid #e2e8f0; font-size: 14px;">Assigned Group Domain</td>
          <td style="padding: 14px 18px; font-weight: 600; color: #1e293b; border-bottom: 1px solid #e2e8f0; font-size: 14px;">${escapeHtml(groupInfo.groupName)}</td>
        </tr>
        <tr>
          <td style="padding: 14px 18px; background-color: #f8fafc; font-weight: 600; color: #475569; border-bottom: 1px solid #e2e8f0; font-size: 14px;">Senior TL Coordinates</td>
          <td style="padding: 14px 18px; border-bottom: 1px solid #e2e8f0; font-size: 14px;"><a href="mailto:${groupInfo.email}" style="color: #0a192f; text-decoration: none; font-weight: 600;">${escapeHtml(groupInfo.email)}</a></td>
        </tr>
        <tr>
          <td style="padding: 14px 18px; background-color: #f8fafc; font-weight: 600; color: #475569; font-size: 14px;">Communications Link</td>
          <td style="padding: 14px 18px; font-size: 14px;"><a href="${groupInfo.whatsapp}" target="_blank" style="color: #16a34a; text-decoration: none; font-weight: 700;">Secure Group Link &rarr;</a></td>
        </tr>
      </table>
      ${getCommonFooter()}
    </div>
  `; 
}

function srTlSummaryTemplate(groupName, rowsHtml, total) {
  return `
    <div style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; color: #1e293b; line-height: 1.6; width: 100%; border: 1px solid #e2e8f0; padding: 30px; border-radius: 12px; box-shadow: 0 4px 6px -1px rgba(0,0,0,0.05);">
      <div style="background: linear-gradient(135deg, #0a192f 0%, #1e293b 100%); color: #ffffff; padding: 24px; text-align: center; border-radius: 8px; font-size: 18px; font-weight: 600; letter-spacing: 0.5px;">
        Cohort Deployment Log: ${escapeHtml(groupName)}
      </div>
      <p style="margin-top: 25px; font-size: 15px;">Dear Sr. Team Leader,</p>
      <p style="font-size: 15px; color: #475569;">The UptoSkills HR Management module has processed a fresh human capital deployment batch for your respective division:</p>
      
      <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; padding: 16px; margin-bottom: 25px; border-radius: 8px;">
        <span style="font-size: 12px; text-transform: uppercase; color: #94a3b8; font-weight: 600; letter-spacing: 0.5px; display: block; margin-bottom: 4px;">Division Assignment Mappings</span>
        <strong style="font-size: 16px; color: #0a192f;">${escapeHtml(groupName)}</strong>
        <span style="color: #cbd5e1; margin: 0 10px;">|</span>
        <span style="font-size: 14px; color: #475569;">Total Ingested Records: <strong style="color: #16a34a;">${total}</strong></span>
      </div>
      
      <div style="width: 100%; overflow-x: auto;">
        <table style="width: 100%; border-collapse: separate; border-spacing: 0; font-size: 13px; border: 1px solid #e2e8f0; border-radius: 8px; overflow: hidden;">
          <thead>
            <tr style="background-color: #f8fafc; color: #475569; text-align: left; font-weight: 600;">
              <th style="padding: 12px; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;">Date</th>
              <th style="padding: 12px; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;">Intern Code</th>
              <th style="padding: 12px; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;">Full Name</th>
              <th style="padding: 12px; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;">Email ID</th>
              <th style="padding: 12px; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;">Mobile No.</th>
              <th style="padding: 12px; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;">Domain</th>
              <th style="padding: 12px; border-bottom: 1px solid #e2e8f0; border-right: 1px solid #e2e8f0;">Start Date</th>
              <th style="padding: 12px; border-bottom: 1px solid #e2e8f0;">End Date</th>
            </tr>
          </thead>
          <tbody style="color: #333333;">
            ${rowsHtml}
          </tbody>
        </table>
      </div>
      ${getCommonFooter()}
    </div>
  `; 
}
