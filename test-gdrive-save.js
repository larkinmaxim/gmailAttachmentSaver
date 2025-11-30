/**
 * Clean 2-Stage OAuth Connection Test
 * 
 * Purpose: Test the 2-stage authentication flow for PMO webhook integration
 * 
 * Usage: Run testClean2StageOAuth() function
 * 
 * This test will:
 * - Stage 1: Get OAuth token from Trimble auth server
 * - Stage 2: Call PMO webhook with Bearer token
 * - Display all environment variables for validation
 */

// ===== ENVIRONMENT VARIABLE GETTER FUNCTIONS =====

/**
 * Get Trimble Auth Server Token from Script Properties
 */
function getTrimbleAuthServerToken() {
  var properties = PropertiesService.getScriptProperties();
  var token = properties.getProperty('TRIMBLE_AUTH_SERVER_TOKEN');
  
  if (!token || token === 'your_auth_server_token_here') {
    throw new Error('TRIMBLE_AUTH_SERVER_TOKEN not configured. Please set your auth server token.');
  }
  
  return token;
}

/**
 * Get PMO Webhook URL from Script Properties
 */
function getPMOWebhookURL() {
  var properties = PropertiesService.getScriptProperties();
  var url = properties.getProperty('PMO_WEBHOOK_URL');
  
  if (!url) {
    throw new Error('PMO_WEBHOOK_URL not configured.');
  }
  
  return url;
}

/**
 * Get Trimble OAuth URL from Script Properties
 */
function getTrimbleOAuthURL() {
  var properties = PropertiesService.getScriptProperties();
  var url = properties.getProperty('TRIMBLE_OAUTH_URL');
  
  if (!url) {
    return 'https://stage.id.trimblecloud.com/oauth/token';
  }
  
  return url;
}

/**
 * Get OAuth Grant Type from Script Properties
 */
function getOAuthGrantType() {
  var properties = PropertiesService.getScriptProperties();
  var grantType = properties.getProperty('OAUTH_GRANT_TYPE');
  
  if (!grantType) {
    return 'client_credentials';
  }
  
  return grantType;
}

/**
 * Get OAuth Scope from Script Properties
 */
function getOAuthScope() {
  var properties = PropertiesService.getScriptProperties();
  var scope = properties.getProperty('OAUTH_SCOPE');
  
  if (!scope) {
    return 'Agentic-N8N-Webhook';
  }
  
  return scope;
}

// ===== CLEAN 2-STAGE OAUTH TEST =====

/**
 * Clean hardcoded 2-stage OAuth connection test
 * Shows all environment variables and tests the complete flow
 */
function testClean2StageOAuth() {
  console.log("=".repeat(80));
  console.log("    CLEAN 2-STAGE OAUTH CONNECTION TEST");
  console.log("=".repeat(80));
  console.log("Test started:", new Date().toISOString());
  
  var testResults = {
    stage1: null,
    stage2: null,
    overall: false,
    variables: {},
    errors: []
  };
  
  try {
    // === DISPLAY ALL ENVIRONMENT VARIABLES ===
    console.log("\n📋 ENVIRONMENT VARIABLES VALIDATION:");
    console.log("=".repeat(50));
    
    try {
      testResults.variables.authServerToken = getTrimbleAuthServerToken();
      console.log("✅ TRIMBLE_AUTH_SERVER_TOKEN: " + testResults.variables.authServerToken.substring(0, 20) + "... (" + testResults.variables.authServerToken.length + " chars)");
    } catch (error) {
      console.log("❌ TRIMBLE_AUTH_SERVER_TOKEN: " + error.message);
      testResults.errors.push("Auth server token: " + error.message);
    }
    
    try {
      testResults.variables.oauthUrl = getTrimbleOAuthURL();
      console.log("✅ TRIMBLE_OAUTH_URL: " + testResults.variables.oauthUrl);
    } catch (error) {
      console.log("❌ TRIMBLE_OAUTH_URL: " + error.message);
      testResults.errors.push("OAuth URL: " + error.message);
    }
    
    try {
      testResults.variables.webhookUrl = getPMOWebhookURL();
      console.log("✅ PMO_WEBHOOK_URL: " + testResults.variables.webhookUrl);
    } catch (error) {
      console.log("❌ PMO_WEBHOOK_URL: " + error.message);
      testResults.errors.push("Webhook URL: " + error.message);
    }
    
    try {
      testResults.variables.grantType = getOAuthGrantType();
      console.log("✅ OAUTH_GRANT_TYPE: " + testResults.variables.grantType);
    } catch (error) {
      console.log("❌ OAUTH_GRANT_TYPE: " + error.message);
      testResults.errors.push("Grant type: " + error.message);
    }
    
    try {
      testResults.variables.scope = getOAuthScope();
      console.log("✅ OAUTH_SCOPE: " + testResults.variables.scope);
    } catch (error) {
      console.log("❌ OAUTH_SCOPE: " + error.message);
      testResults.errors.push("OAuth scope: " + error.message);
    }
    
    // Stop if we have configuration errors
    if (testResults.errors.length > 0) {
      console.log("\n❌ CONFIGURATION ERRORS - Cannot proceed with test");
      console.log("Fix the following issues:");
      for (var i = 0; i < testResults.errors.length; i++) {
        console.log("  " + (i + 1) + ". " + testResults.errors[i]);
      }
      return testResults;
    }
    
    // === STAGE 1: OAUTH TOKEN RETRIEVAL ===
    console.log("\n🔐 STAGE 1: OAuth Token Retrieval");
    console.log("=".repeat(50));
    
    var stage1Start = new Date().getTime();
    
    console.log("Making OAuth request:");
    console.log("  URL: " + testResults.variables.oauthUrl);
    console.log("  Method: POST");
    console.log("  Authorization: " + testResults.variables.authServerToken.substring(0, 15) + "...");
    console.log("  Grant Type: " + testResults.variables.grantType);
    console.log("  Scope: " + testResults.variables.scope);
    
    var payload = 'grant_type=' + encodeURIComponent(testResults.variables.grantType) + 
                  '&scope=' + encodeURIComponent(testResults.variables.scope);
    
    console.log("  Payload: " + payload);
    
    var stage1Response = UrlFetchApp.fetch(testResults.variables.oauthUrl, {
      method: 'POST',
      headers: {
        'Authorization': testResults.variables.authServerToken,
        'Content-Type': 'application/x-www-form-urlencoded'
      },
      payload: payload,
      muteHttpExceptions: true,
      timeout: 15000
    });
    
    var stage1Duration = new Date().getTime() - stage1Start;
    var stage1Code = stage1Response.getResponseCode();
    var stage1Text = stage1Response.getContentText();
    
    console.log("\nStage 1 Results:");
    console.log("  Duration: " + stage1Duration + "ms");
    console.log("  Response Code: " + stage1Code);
    console.log("  Response Length: " + stage1Text.length + " chars");
    
    if (stage1Code === 200) {
      try {
        var oauthData = JSON.parse(stage1Text);
        var accessToken = oauthData.access_token;
        var expiresIn = oauthData.expires_in || 3600;
        var tokenType = oauthData.token_type || 'Bearer';
        
        console.log("✅ Stage 1 SUCCESS:");
        console.log("  Access Token: " + accessToken.substring(0, 30) + "... (truncated)");
        console.log("  Token Type: " + tokenType);
        console.log("  Expires In: " + expiresIn + " seconds (" + Math.floor(expiresIn/60) + " minutes)");
        
        testResults.stage1 = {
          success: true,
          accessToken: accessToken,
          tokenType: tokenType,
          expiresIn: expiresIn,
          duration: stage1Duration
        };
        
        // === STAGE 2: PMO WEBHOOK CALL ===
        console.log("\n🌐 STAGE 2: PMO Webhook Call with Bearer Token");
        console.log("=".repeat(50));
        
        var stage2Start = new Date().getTime();
        var testTicket = 'CXPRODELIVERY-7081';
        
        console.log("Making webhook request:");
        console.log("  URL: " + testResults.variables.webhookUrl);
        console.log("  Method: POST");
        console.log("  Authorization: Bearer " + accessToken.substring(0, 30) + "...");
        console.log("  Test Ticket: " + testTicket);
        
        var webhookPayload = JSON.stringify({"text": testTicket});
        console.log("  Payload: " + webhookPayload);
        
        var stage2Response = UrlFetchApp.fetch(testResults.variables.webhookUrl, {
          method: 'POST',
          headers: {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer ' + accessToken
          },
          payload: webhookPayload,
          muteHttpExceptions: true,
          timeout: 15000
        });
        
        var stage2Duration = new Date().getTime() - stage2Start;
        var stage2Code = stage2Response.getResponseCode();
        var stage2Text = stage2Response.getContentText();
        
        console.log("\nStage 2 Results:");
        console.log("  Duration: " + stage2Duration + "ms");
        console.log("  Response Code: " + stage2Code);
        console.log("  Response Length: " + stage2Text.length + " chars");
        
        if (stage2Code === 200) {
          console.log("✅ Stage 2 SUCCESS:");
          console.log("  Raw Response: " + stage2Text);
          
          testResults.stage2 = {
            success: true,
            responseCode: stage2Code,
            response: stage2Text,
            duration: stage2Duration,
            testTicket: testTicket
          };
          
        } else {
          console.log("❌ Stage 2 FAILED:");
          console.log("  HTTP Status: " + stage2Code);
          console.log("  Error Response: " + stage2Text);
          
          testResults.stage2 = {
            success: false,
            responseCode: stage2Code,
            error: stage2Text,
            duration: stage2Duration
          };
          
          // Analyze Stage 2 errors
          if (stage2Code === 401) {
            console.log("💡 401 Unauthorized - Token may be invalid or expired");
          } else if (stage2Code === 403) {
            console.log("💡 403 Forbidden - Token lacks required permissions");
          } else if (stage2Code === 404) {
            console.log("💡 404 Not Found - Webhook URL may be incorrect");
          } else if (stage2Code === 400) {
            console.log("💡 400 Bad Request - Payload format issue");
          }
        }
        
      } catch (stage1ParseError) {
        console.log("❌ Stage 1 JSON parsing failed:");
        console.log("  Parse Error: " + stage1ParseError.message);
        console.log("  Raw Response: " + stage1Text);
        
        testResults.stage1 = {
          success: false,
          error: "JSON parsing failed: " + stage1ParseError.message,
          duration: stage1Duration
        };
      }
      
    } else {
      console.log("❌ Stage 1 FAILED:");
      console.log("  HTTP Status: " + stage1Code);
      console.log("  Error Response: " + stage1Text);
      
      testResults.stage1 = {
        success: false,
        responseCode: stage1Code,
        error: stage1Text,
        duration: stage1Duration
      };
      
      // Analyze Stage 1 errors
      if (stage1Code === 401) {
        console.log("💡 401 Unauthorized - Invalid auth server token");
      } else if (stage1Code === 400) {
        console.log("💡 400 Bad Request - Invalid grant_type or scope");
      } else if (stage1Code === 403) {
        console.log("💡 403 Forbidden - Token lacks OAuth permissions");
      }
    }
    
    // Calculate overall result
    testResults.overall = testResults.stage1 && testResults.stage1.success && 
                         testResults.stage2 && testResults.stage2.success;
    
    var totalDuration = (testResults.stage1 ? testResults.stage1.duration : 0) + 
                       (testResults.stage2 ? testResults.stage2.duration : 0);
    
    // === FINAL SUMMARY ===
    console.log("\n" + "=".repeat(80));
    console.log("    CLEAN 2-STAGE OAUTH TEST SUMMARY");
    console.log("=".repeat(80));
    
    console.log("\n📊 STAGE RESULTS:");
    console.log("  Stage 1 (OAuth): " + (testResults.stage1 && testResults.stage1.success ? "✅ SUCCESS" : "❌ FAILED"));
    console.log("  Stage 2 (Webhook): " + (testResults.stage2 && testResults.stage2.success ? "✅ SUCCESS" : "❌ FAILED"));
    
    console.log("\n⏱️ PERFORMANCE:");
    console.log("  Stage 1 Duration: " + (testResults.stage1 ? testResults.stage1.duration : 0) + "ms");
    console.log("  Stage 2 Duration: " + (testResults.stage2 ? testResults.stage2.duration : 0) + "ms");
    console.log("  Total Duration: " + totalDuration + "ms");
    
    console.log("\n🎯 OVERALL RESULT: " + (testResults.overall ? "🎉 2-STAGE OAUTH FLOW WORKING!" : "❌ 2-STAGE OAUTH FLOW FAILED"));
    
    if (testResults.overall) {
      console.log("\n✅ SUCCESS! Both stages completed successfully");
      console.log("  • OAuth token retrieved from Trimble auth server");
      console.log("  • PMO webhook responded successfully with Bearer token");
      console.log("  • 2-stage authentication flow is fully functional");
    } else {
      console.log("\n❌ FAILURE! One or both stages failed");
      if (!testResults.stage1 || !testResults.stage1.success) {
        console.log("  • Fix Stage 1 OAuth token issues first");
        console.log("  • Verify TRIMBLE_AUTH_SERVER_TOKEN is correct");
        console.log("  • Check token format and permissions");
      }
      if (testResults.stage1 && testResults.stage1.success && (!testResults.stage2 || !testResults.stage2.success)) {
        console.log("  • OAuth working but webhook failing");
        console.log("  • Check webhook URL and token permissions");
        console.log("  • Verify token has webhook access scope");
      }
    }
    
    console.log("\n💡 CURL EQUIVALENTS:");
    console.log("# Stage 1 - Get OAuth Token:");
    console.log("curl --location '" + testResults.variables.oauthUrl + "' \\");
    console.log("--header 'Authorization: TRIMBLE_AUTH_SERVER_TOKEN' \\");  
    console.log("--header 'Content-Type: application/x-www-form-urlencoded' \\");
    console.log("--data-urlencode 'grant_type=" + testResults.variables.grantType + "' \\");
    console.log("--data-urlencode 'scope=" + testResults.variables.scope + "'");
    
    console.log("\n# Stage 2 - Use Bearer Token:");
    console.log("curl --location '" + testResults.variables.webhookUrl + "' \\");
    console.log("--header 'Content-Type: application/json' \\");
    console.log("--header 'Authorization: Bearer CACHED_OAUTH_TOKEN' \\");
    console.log("--data '{\"text\": \"CXPRODELIVERY-7081\"}'");
    
    return testResults;
    
  } catch (error) {
    console.error("❌ Test failed with exception:", error.message);
    console.error("Stack trace:", error.stack);
    
    return {
      success: false,
      error: error.message,
      stage1: null,
      stage2: null,
      overall: false,
      variables: testResults.variables || {}
    };
  }
}

/**
 * Quick variable check - just display all environment variables
 */
function showAllVariables() {
  console.log("=== ENVIRONMENT VARIABLES CHECK ===");
  console.log("Timestamp:", new Date().toISOString());
  
  try {
    console.log("TRIMBLE_AUTH_SERVER_TOKEN: " + getTrimbleAuthServerToken().substring(0, 20) + "... (" + getTrimbleAuthServerToken().length + " chars)");
  } catch (e) {
    console.log("TRIMBLE_AUTH_SERVER_TOKEN: ERROR - " + e.message);
  }
  
  try {
    console.log("TRIMBLE_OAUTH_URL: " + getTrimbleOAuthURL());
  } catch (e) {
    console.log("TRIMBLE_OAUTH_URL: ERROR - " + e.message);
  }
  
  try {
    console.log("PMO_WEBHOOK_URL: " + getPMOWebhookURL());
  } catch (e) {
    console.log("PMO_WEBHOOK_URL: ERROR - " + e.message);
  }
  
  try {
    console.log("OAUTH_GRANT_TYPE: " + getOAuthGrantType());
  } catch (e) {
    console.log("OAUTH_GRANT_TYPE: ERROR - " + e.message);
  }
  
  try {
    console.log("OAUTH_SCOPE: " + getOAuthScope());
  } catch (e) {
    console.log("OAUTH_SCOPE: ERROR - " + e.message);
  }
  
  console.log("=== END VARIABLES CHECK ===");
}