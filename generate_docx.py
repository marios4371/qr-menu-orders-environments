#!/usr/bin/env python3
"""Generate comprehensive project documentation as .docx"""

from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml.ns import qn
import datetime

doc = Document()

# ─── Styles ───
style = doc.styles['Normal']
style.font.name = 'Calibri'
style.font.size = Pt(11)
style.paragraph_format.space_after = Pt(6)
style.paragraph_format.line_spacing = 1.15

for level in range(1, 5):
    h = doc.styles[f'Heading {level}']
    h.font.name = 'Calibri'
    h.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

# Code style
code_style = doc.styles.add_style('CodeBlock', WD_STYLE_TYPE.PARAGRAPH)
code_style.font.name = 'Consolas'
code_style.font.size = Pt(9)
code_style.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
code_style.paragraph_format.space_before = Pt(4)
code_style.paragraph_format.space_after = Pt(4)
code_style.paragraph_format.left_indent = Cm(1)

def add_code(text):
    for line in text.strip().split('\n'):
        doc.add_paragraph(line, style='CodeBlock')

def add_bullet(text, bold_prefix=None):
    p = doc.add_paragraph(style='List Bullet')
    if bold_prefix:
        run = p.add_run(bold_prefix)
        run.bold = True
        p.add_run(text)
    else:
        p.add_run(text)

def add_note(text):
    p = doc.add_paragraph()
    run = p.add_run(text)
    run.italic = True
    run.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    run.font.size = Pt(10)

# ═══════════════════════════════════════════════════
# COVER PAGE
# ═══════════════════════════════════════════════════
for _ in range(6):
    doc.add_paragraph()

title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('QR Menu & Restaurant Management Platform')
run.bold = True
run.font.size = Pt(26)
run.font.color.rgb = RGBColor(0x1A, 0x1A, 0x2E)

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('Complete Technical Documentation & System Architecture')
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

doc.add_paragraph()
meta = doc.add_paragraph()
meta.alignment = WD_ALIGN_PARAGRAPH.CENTER
meta.add_run(f'Version 1.0 — {datetime.date.today().strftime("%B %d, %Y")}').font.size = Pt(11)

meta2 = doc.add_paragraph()
meta2.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = meta2.add_run('Serverless B2B2C Platform — Hospitality Industry')
run.font.size = Pt(11)
run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

doc.add_page_break()

# ═══════════════════════════════════════════════════
# TABLE OF CONTENTS (placeholder)
# ═══════════════════════════════════════════════════
doc.add_heading('Table of Contents', level=1)
toc_items = [
    '1. Product Vision & Scope',
    '2. User Roles & Experience Flows',
    '3. High-Level System Architecture',
    '4. AWS Cloud Infrastructure',
    '5. Database Schema (DynamoDB)',
    '6. Backend API — Java 17 Lambda Handlers',
    '7. Customer Web Frontend',
    '8. Ordering & Kitchen Display Web Apps',
    '9. Business Owner Mobile Application (MenuAdmin)',
    '10. Infrastructure as Code (Terraform)',
    '11. CI/CD Pipelines (GitHub Actions)',
    '12. Data Flow & Component Interaction',
    '13. Epic & Kanban Roadmap',
]
for item in toc_items:
    doc.add_paragraph(item, style='List Number')
doc.add_page_break()

# ═══════════════════════════════════════════════════
# SECTION 1: Product Vision & Scope
# ═══════════════════════════════════════════════════
doc.add_heading('1. Product Vision & Scope', level=1)
add_note('Notion Source: 00_Master_Business_Plan, 01_Product_Vision_And_Roles')

doc.add_paragraph(
    'This project is a B2B2C Serverless Microservices platform designed for the hospitality industry. '
    'It bridges the gap between customer convenience and restaurant operational efficiency. '
    'Customers scan a QR code at their table to view a digital menu and place orders directly from their mobile browser. '
    'Restaurant staff use a web-based dashboard (staff.html) to view and process orders at their respective stations '
    '(Bar, Kitchen). Meanwhile, the business owner manages the restaurant through a dedicated React Native mobile app (MenuAdmin).'
)

doc.add_paragraph(
    'The platform goes beyond a simple QR menu viewer — it functions as a lightweight POS/RMS (Point of Sale / Restaurant '
    'Management System) with intelligent order routing, station-based displays, and a dedicated business management app for the owner.'
)

doc.add_heading('1.1 Core Capabilities', level=2)
add_bullet('Multi-tenant architecture supporting multiple restaurant brands (e.g. Nissos, Rakoumel, Theristis)')
add_bullet('QR-code-driven customer ordering with zero app installation')
add_bullet('Real-time order routing to Bar and Kitchen preparation stations')
add_bullet('Table claiming and shared-pool order assignment for waiters')
add_bullet('Full order lifecycle management (NEW -> CLAIMED -> PARTIAL -> READY -> CLOSED)')
add_bullet('Dynamic menu management via dedicated mobile app for the business owner (MenuAdmin)')
add_bullet('Web-based staff dashboard (staff.html) for preparation stations')
add_bullet('Planned business portfolio: analytics, income/expenses, inventory management')

# ═══════════════════════════════════════════════════
# SECTION 2: User Roles
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('2. User Roles & Experience Flows', level=1)
add_note('Notion Source: 01_Product_Vision_And_Roles')

doc.add_heading('2.1 Customer (Web)', level=2)
doc.add_paragraph(
    'The customer scans a QR code placed at their table. This opens a static web application in their mobile browser '
    'showing the shop-specific menu. When they are ready to order, they press the "START ORDER" button embedded in the menu UI. '
    'This redirects them to app.html — a "chameleon" ordering page that inherits UI elements (theme, colors, branding) '
    'from the currently running menu. There, they enter their table number, build their order, and submit it. '
    'The order is sent as a POST request to the Java order-service API Gateway (SAM-generated), which invokes the '
    'CreateOrderHandler Lambda function. The order source is tagged as CUSTOMER_QR.'
)

doc.add_heading('2.2 Staff (Web — staff.html)', level=2)
doc.add_paragraph(
    'Staff members use a web-based dashboard (staff.html in qr-menu-orders-environments), NOT the React Native mobile app. '
    'The dashboard is accessed via a browser and is designed for display screens at preparation stations.'
)

doc.add_heading('2.2.1 Station: BAR & KITCHEN', level=3)
doc.add_paragraph(
    'Each preparation station (Bar, Kitchen) opens staff.html filtered by their station. '
    'They see only the orders relevant to them — the Bar sees items where station equals "BAR", '
    'the Kitchen sees items where station equals "KITCHEN". '
    'When an item is prepared, they press "MARK AS DONE", which sends a PATCH request to UpdateOrderStatusHandler. '
    'If all items across all stations are done, the parent order transitions to READY.'
)

doc.add_heading('2.2.2 Waiter / Service Role', level=3)
doc.add_paragraph(
    'The waiter role (planned) will also use the web-based dashboard with the following features:'
)
add_bullet('All NEW, unassigned customer orders waiting to be claimed.', 'Shared Pool (Dexameni): ')
add_bullet('The waiter presses "ANALIPSI" (Claim) to take ownership of an order. This sets claimedBy to their name and status to CLAIMED.', 'Claiming: ')
add_bullet('A grid view of tables — Green (my tables), Red (other waiters\' tables), Grey (empty).', 'Floor Plan: ')
add_bullet('Can create new orders directly (source: WAITER_PDA), append items to existing active tables, and mark tables as Paid/Closed.', 'Actions: ')

doc.add_heading('2.3 Business Owner (React Native App — MenuAdmin)', level=2)
doc.add_paragraph(
    'The business owner uses the React Native mobile app (MenuAdmin). Staff members do NOT have access to this application. '
    'Currently, the app provides full menu management (CRUD for categories, products, prices, descriptions). '
    'It is planned to evolve into a complete business portfolio with the following features:'
)
add_bullet('Statistics on orders, peak hours, average order value', 'Order Analytics: ')
add_bullet('Most popular items, best sellers by category', 'Top Products: ')
add_bullet('Revenue tracking, cost management, profit margins', 'Income & Expenses: ')
add_bullet('Stock management for the restaurant, cafe, and bar. Track supplies, set low-stock alerts.', 'Inventory / Storage (Kava): ')
add_bullet('A comprehensive system for managing all operational needs of the business', 'General Business Organization: ')

# ═══════════════════════════════════════════════════
# SECTION 3: High-Level Architecture
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('3. High-Level System Architecture', level=1)
add_note('Notion Source: 00_Master_Business_Plan, 02_Cloud_Architecture_And_DevOps')

doc.add_paragraph(
    'The platform operates on a 100% serverless paradigm on AWS, ensuring high scalability, zero idle-server costs, '
    'and rapid deployments. The architecture is composed of five major layers:'
)

# Architecture table
table = doc.add_table(rows=7, cols=3)
table.style = 'Light Shading Accent 1'
headers = ['Layer', 'Technology', 'Repository']
for i, h in enumerate(headers):
    table.rows[0].cells[i].text = h
    for paragraph in table.rows[0].cells[i].paragraphs:
        for run in paragraph.runs:
            run.bold = True

data = [
    ['Customer Frontend (Web)', 'HTML5, CSS3, Vanilla JS — Static S3/CloudFront', 'menu-frontend-CICD-Automation'],
    ['Ordering & Kitchen Web Apps', 'HTML5, JS — Hosted on S3', 'qr-menu-orders-environments'],
    ['Business Owner App (Mobile)', 'React Native, Expo, TypeScript/JS', 'MenuAdmin'],
    ['Menu Management Backend', 'Node.js 20, Lambda Function URL, Terraform', 'menu-backend-CICD-Automation'],
    ['Order Service Backend', 'Java 17, AWS SAM, API Gateway + Lambda', 'qr-menu-order-service'],
]
for row_idx, row_data in enumerate(data, 1):
    for col_idx, val in enumerate(row_data):
        table.rows[row_idx].cells[col_idx].text = val

doc.add_paragraph()
doc.add_heading('3.1 Component Interaction Overview', level=2)
doc.add_paragraph(
    'The flow works as follows: (1) The Customer opens the static menu website hosted on S3/CloudFront. '
    '(2) The Node.js Lambda fetches shop data from DynamoDB and injects it into the HTML as window.SHOP_DATA before serving. '
    '(3) They press "START ORDER" and are redirected to app.html, a chameleon page that inherits the menu\'s UI theme. '
    '(4) Upon order submission, a POST hits CreateOrderHandler, which writes to DynamoDB with status=NEW. '
    '(5) The staff dashboard (staff.html) polls GetOrdersHandler and displays orders per station. '
    '(6) Bar/Kitchen stations see their filtered items and mark them done (action=ITEM_DONE). '
    '(7) When all items are done, the order becomes READY and the waiter is notified. '
    '(8) The waiter closes the order (action=CLOSE) after payment. '
    'Separately, the business owner uses the MenuAdmin React Native app to manage the menu and (in the future) '
    'view analytics, track income/expenses, and manage inventory.'
)

doc.add_paragraph(
    'The two backend systems are independent: The Node.js Lambda ("qr-menu") handles menu management, owner authentication, '
    'and S3 template serving via a Lambda Function URL (no API Gateway). The Java Lambdas (SAM) handle order processing '
    'via a SAM-generated API Gateway. Both share the same DynamoDB "Menus" table.'
)

# ═══════════════════════════════════════════════════
# SECTION 4: AWS Cloud Infrastructure
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('4. AWS Cloud Infrastructure', level=1)
add_note('Notion Source: 02_Cloud_Architecture_And_DevOps')

doc.add_heading('4.1 Core AWS Services', level=2)
add_bullet('API Gateway (Java/SAM only): SAM-generated ServerlessRestApi that routes order-service endpoints (/orders, /menu, /orders/status). Handles CORS. Used ONLY by the Java order-service, NOT by the Node.js Lambda.', 'API Gateway: ')
add_bullet('Lambda Function URL (Node.js): The Node.js Lambda ("qr-menu") is invoked directly via a Lambda Function URL, bypassing API Gateway entirely. This is used by the customer template serving and the MenuAdmin app.', 'Lambda Function URL: ')
add_bullet('AWS Lambda — Java 17 (SAM): Simple Java with no framework. Four handler functions deployed as individual Lambda resources via SAM template.yaml. Deployed via AWS SAM CLI (sam build / sam deploy).', 'AWS Lambda (Java): ')
add_bullet('AWS Lambda — Node.js 20 (Terraform): The menu management Lambda (index.mjs + db.mjs). Deployed via Terraform. Handles owner auth, menu CRUD, and S3 template serving.', 'AWS Lambda (Node.js): ')
add_bullet('Amazon DynamoDB: Two tables — "Menus" (menu configurations per shop) and "orders" (order documents with lifecycle).', 'DynamoDB: ')
add_bullet('Amazon S3: Hosts static web client files (per-shop HTML/CSS/JS templates) and Terraform state. Bucket: qr-templates-io.', 'Amazon S3: ')
add_bullet('Amazon CloudFront: CDN distribution for the S3-hosted customer frontend.', 'CloudFront: ')

doc.add_heading('4.2 Region & Deployment', level=2)
doc.add_paragraph('All resources are deployed in the eu-central-1 (Frankfurt) AWS region.')

# ═══════════════════════════════════════════════════
# SECTION 5: Database Schema
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('5. Database Schema (DynamoDB)', level=1)
add_note('Notion Source: 03_Database_Schema_DynamoDB — This is the Single Source of Truth for the Data Model.')

doc.add_heading('5.1 Table: "Menus"', level=2)
doc.add_paragraph(
    'Stores the full configuration of each restaurant/shop, including its menu categories, products, '
    'settings (logo, theme), and admin password. Each product must specify a "station" field (BAR or KITCHEN) '
    'to enable the smart routing system.'
)
add_bullet('Partition Key: shop_id (String) — e.g. "nissos", "rakoumel"', 'Primary Key: ')
add_bullet('Contains: menu (array of categories with items), settings, features, theme, password')

doc.add_heading('5.2 Table: "orders"', level=2)
doc.add_paragraph('Stores all order documents. Uses a composite key design for efficient per-shop querying.')

doc.add_heading('5.2.1 Order Object (Root Level)', level=3)

order_table = doc.add_table(rows=10, cols=3)
order_table.style = 'Light Shading Accent 1'
for i, h in enumerate(['Field', 'Type', 'Description']):
    order_table.rows[0].cells[i].text = h
    for run in order_table.rows[0].cells[i].paragraphs[0].runs:
        run.bold = True

order_fields = [
    ['shopId', 'String (PK)', 'Partition Key — Links to the restaurant'],
    ['orderId', 'String (SK)', 'Sort Key — Format: ORDER#TIMESTAMP#UUID'],
    ['status', 'String', 'NEW | CLAIMED | PARTIAL | READY | CLOSED'],
    ['paymentStatus', 'String', 'UNPAID | PAID'],
    ['claimedBy', 'String', 'Waiter\'s name or "NONE"'],
    ['source', 'String', 'CUSTOMER_QR | WAITER_PDA'],
    ['tableNumber', 'String', 'The physical table identifier'],
    ['totalAmount', 'Number', 'Sum of all item prices'],
    ['createdAt', 'String', 'ISO 8601 timestamp'],
]
for row_idx, row_data in enumerate(order_fields, 1):
    for col_idx, val in enumerate(row_data):
        order_table.rows[row_idx].cells[col_idx].text = val

doc.add_paragraph()
doc.add_heading('5.2.2 OrderItem Object (Inside "items" array)', level=3)

item_table = doc.add_table(rows=7, cols=3)
item_table.style = 'Light Shading Accent 1'
for i, h in enumerate(['Field', 'Type', 'Description']):
    item_table.rows[0].cells[i].text = h
    for run in item_table.rows[0].cells[i].paragraphs[0].runs:
        run.bold = True

item_fields = [
    ['productId', 'String', 'References the product in the Menus table'],
    ['name', 'String', 'Human-readable product name'],
    ['quantity', 'Number', 'Units ordered'],
    ['station', 'String', 'BAR | KITCHEN — Determines routing'],
    ['itemStatus', 'String', 'PENDING | DONE'],
    ['timestamp', 'String', 'Differentiates appended items'],
]
for row_idx, row_data in enumerate(item_fields, 1):
    for col_idx, val in enumerate(row_data):
        item_table.rows[row_idx].cells[col_idx].text = val

doc.add_paragraph()
doc.add_heading('5.2.3 Order Lifecycle', level=3)
doc.add_paragraph(
    'An order goes through the following state transitions:'
)
add_bullet('NEW: Customer submitted or Waiter created. Waiting in the shared pool.')
add_bullet('CLAIMED: A waiter claimed the order. claimedBy is set. Bar/Kitchen can now see the items.')
add_bullet('PARTIAL: One station (e.g. Bar) finished all its items, but another (Kitchen) hasn\'t. '
           'This is an intermediate state.')
add_bullet('READY: All items across all stations have itemStatus=DONE. The waiter is notified to deliver.')
add_bullet('CLOSED: Payment received. paymentStatus=PAID. Order archived.')

# ═══════════════════════════════════════════════════
# SECTION 6: Backend API
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('6. Backend API — Java 17 Lambda Handlers', level=1)
add_note('Notion Source: 04_Backend_API_Java — Repository: qr-menu-order-service')

doc.add_paragraph(
    'The order-processing backend is built with Java 17 using the AWS Serverless Application Model (SAM). '
    'Each endpoint is a separate Lambda function, defined in template.yaml, that receives APIGatewayProxyRequestEvent '
    'objects and returns APIGatewayProxyResponseEvent. They share common patterns: static DynamoDB client initialization '
    'for cold-start optimization, Gson for JSON serialization, and a standardized createResponse() helper with CORS headers.'
)

doc.add_heading('6.1 CreateOrderHandler — POST /orders', level=2)
add_note('File: qr-menu-order-service/src/main/java/com/qrmenu/orders/handlers/CreateOrderHandler.java')
doc.add_paragraph(
    'Receives a new order from the customer web app (source=CUSTOMER_QR) or from a waiter\'s PDA (source=WAITER_PDA). '
    'The handler deserializes the JSON body into an Order object, validates that shopId exists, then sets default state values: '
    'status="NEW", paymentStatus="UNPAID", claimedBy="NONE". If source is not provided, it defaults to "CUSTOMER_QR". '
    'It loops through all order items and sets itemStatus="PENDING" and timestamp=current ISO timestamp. '
    'Station validation: if any item has a null/empty station field, it defaults to "KITCHEN" as a fail-safe. '
    'Similarly, if productId is null, the item name is used as a fallback ID. '
    'The handler then generates a composite orderId (ORDER#<ISO_TIMESTAMP>#<UUID>), maps all fields to DynamoDB '
    'AttributeValue objects, and stores the items array as a JSON string. Returns 201 with the generated orderId.'
)

doc.add_paragraph('Key code reference — Default state initialization and station validation:')
add_code('''// Default state values
newOrder.setStatus("NEW");
newOrder.setPaymentStatus("UNPAID");
newOrder.setClaimedBy("NONE");
if (newOrder.getSource() == null) {
    newOrder.setSource("CUSTOMER_QR");
}

// Per-item initialization
for (OrderItem item : newOrder.getOrderItem()) {
    item.setItemStatus("PENDING");
    item.setTimestamp(currentTimestamp);
    if (item.getStation() == null || item.getStation().trim().isEmpty()) {
        item.setStation("KITCHEN"); // Fail-safe default
    }
}''')

doc.add_heading('6.2 GetOrdersHandler — GET /orders?shopId={id}', level=2)
add_note('File: qr-menu-order-service/src/main/java/com/qrmenu/orders/handlers/GetOrdersHandler.java')
doc.add_paragraph(
    'Fetches all orders for a given shop. Uses a DynamoDB Query with the Partition Key (shopId) and a '
    'begins_with condition on the Sort Key (orderId starts with "ORDER#") to efficiently retrieve only order records. '
    'The handler maps DynamoDB items back to plain Java Maps and deserializes the stored JSON items string back to '
    'an object array for the response.'
)

doc.add_paragraph('Key code reference — DynamoDB query pattern:')
add_code('''QueryRequest queryRequest = QueryRequest.builder()
    .tableName(TABLE_NAME)
    .keyConditionExpression("#pk = :shopIdValue AND begins_with(#sk, :orderPrefix)")
    .expressionAttributeNames(expressionAttributesNames)
    .expressionAttributeValues(expressionAttributeValues)
    .build();''')

doc.add_heading('6.3 UpdateOrderStatusHandler — PATCH /orders/status', level=2)
add_note('File: qr-menu-order-service/src/main/java/com/qrmenu/orders/handlers/UpdateOrderStatusHandler.java')
doc.add_paragraph(
    'This handler acts as the central "router" for all order state transitions. It receives an OrderUpdateRequest '
    'JSON payload containing an action field and executes the appropriate business logic via a switch statement. '
    'The handler uses a full read-modify-write pattern: it first fetches the current order from DynamoDB (GetItem), '
    'deserializes the items JSON array, applies the mutation, then writes the entire modified item back (PutItem). '
    'This is necessary because actions like ITEM_DONE and APPEND need to read and modify the items array.'
)
add_bullet('CLAIM: Sets status=CLAIMED, claimedBy=employeeName. The waiter takes ownership of the order.', 'ACTION CLAIM: ')
add_bullet('ITEM_DONE: Matches a specific item by productId + itemTimestamp (dual key for disambiguation when '
           'the same product appears multiple times). Sets that item\'s itemStatus=DONE. If ALL items in the order '
           'are now DONE, sets parent status=READY. Otherwise sets status=PARTIAL.', 'ACTION ITEM_DONE: ')
add_bullet('APPEND: Adds newItems (List<OrderItem>) to the existing items array. Resets parent status to CLAIMED '
           'so kitchen/bar sees the new ticket.', 'ACTION APPEND: ')
add_bullet('CLOSE: Sets status=CLOSED, paymentStatus=PAID. Order is archived.', 'ACTION CLOSE: ')

doc.add_paragraph('Action Payload Model (OrderUpdateRequest.java):')
add_code('''{
  "shopId": "nissos",
  "orderId": "ORDER#2026-03-01T10:15:30Z#abc123",
  "action": "CLAIM",        // CLAIM | ITEM_DONE | APPEND | CLOSE
  "employeeName": "George",  // Used by CLAIM
  "productId": "burger_01",  // Used by ITEM_DONE
  "itemTimestamp": "2026-03-01T10:15:30Z",  // Used by ITEM_DONE (dual key)
  "newItems": []              // Used by APPEND
}''')

doc.add_paragraph('Key code reference — Action-based switch routing:')
add_code('''OrderUpdateRequest updateReq = gson.fromJson(request.getBody(), OrderUpdateRequest.class);
// GetItem -> read current state
GetItemResponse getResponse = dynamoDb.getItem(getItemRequest);
List<OrderItem> currentItems = gson.fromJson(currentItemsJson, listType);

switch (action.toUpperCase()) {
    case "CLAIM":
        item.put("status", AttributeValue.builder().s("CLAIMED").build());
        item.put("claimedBy", AttributeValue.builder().s(updateReq.getEmployeeName()).build());
        break;
    case "ITEM_DONE":
        // Match by productId + itemTimestamp
        for (OrderItem oi : currentItems) {
            if (oi.getProductId().equals(updateReq.getProductId()) &&
                oi.getTimestamp().equals(updateReq.getItemTimestamp())) {
                oi.setItemStatus("DONE");
            }
        }
        currentStatus = allDone ? "READY" : "PARTIAL";
        break;
    case "APPEND":
        currentItems.addAll(updateReq.getNewItems());
        item.put("status", AttributeValue.builder().s("CLAIMED").build());
        break;
    case "CLOSE":
        item.put("status", AttributeValue.builder().s("CLOSED").build());
        item.put("paymentStatus", AttributeValue.builder().s("PAID").build());
        break;
}
// PutItem -> write modified state back
dynamoDb.putItem(putRequest);''')

doc.add_heading('6.4 GetMenuHandler — GET /menu?shopId={id}', level=2)
add_note('File: qr-menu-order-service/src/main/java/com/qrmenu/orders/handlers/GetMenuHandler.java')
doc.add_paragraph(
    'Fetches the menu configuration for a specific shop from the "Menus" DynamoDB table. '
    'Uses a DynamoDB GetItem with shop_id as the key. The handler includes a recursive '
    'toPlainObject() utility method that converts complex nested DynamoDB AttributeValue types '
    '(Maps, Lists, Strings, Numbers, Booleans) back into standard Java objects that Gson can serialize to JSON. '
    'Returns the full menu, settings, features, and theme objects.'
)

doc.add_heading('6.5 Data Models', level=2)
add_note('Files: Order.java, OrderItem.java, OrderUpdateRequest.java')
doc.add_paragraph(
    'The Order model contains: shopId, orderId, tableNumber, orderItem (List<OrderItem>), totalAmount, status, '
    'source, createdAt, claimedBy (String — waiter name or "NONE"), and paymentStatus (String — UNPAID or PAID). '
    'The OrderItem model contains: productId, name, quantity, unitPrice, modifiers (List<String>), '
    'station (String — BAR or KITCHEN), itemStatus (String — PENDING or DONE), and timestamp (String — ISO 8601, '
    'used to disambiguate identical products in the same order).'
)
doc.add_paragraph(
    'The OrderUpdateRequest model (new) is the payload for UpdateOrderStatusHandler: shopId, orderId, '
    'action (CLAIM/ITEM_DONE/APPEND/CLOSE), employeeName (for CLAIM), productId + itemTimestamp (for ITEM_DONE), '
    'and newItems (List<OrderItem> for APPEND).'
)

doc.add_heading('6.6 SAM Template (template.yaml)', level=2)
add_note('File: qr-menu-order-service/template.yaml')
doc.add_paragraph(
    'Defines the entire Java backend stack as infrastructure-as-code. Key aspects:'
)
add_bullet('Global CORS configuration: AllowMethods, AllowHeaders, AllowOrigin="*" — applied to all API endpoints')
add_bullet('Creates the "orders" DynamoDB table with PAY_PER_REQUEST billing (shopId as PK, orderId as SK)')
add_bullet('Deploys 4 Lambda functions (CreateOrder, GetOrders, UpdateOrderStatus, GetMenu) using java21 runtime')
add_bullet('Each function gets 512MB memory and 10-15s timeout')
add_bullet('Least-privilege IAM: GetOrders and GetMenu get DynamoDBReadPolicy; CreateOrder and UpdateOrderStatus get '
           'DynamoDBCrudPolicy (UpdateOrderStatus needs both GetItem and PutItem for read-modify-write)')
add_bullet('CodeUri set to "." — SAM builds the project from the repository root')
add_bullet('API events automatically create the API Gateway endpoints with correct HTTP methods and paths')

# ═══════════════════════════════════════════════════
# SECTION 7: Customer Frontend
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('7. Customer Web Frontend', level=1)
add_note('Notion Source: 05_Frontend_Web_Customer — Repository: menu-frontend-CICD-Automation')

doc.add_paragraph(
    'The customer-facing web frontend follows a multi-tenant static hosting architecture. '
    'Each restaurant brand has its own directory under shops/ (e.g., shops/nissos/, shops/rakoumel/, shops/theristis/) '
    'containing a template.html and style.css. These files are deployed to an S3 bucket (qr-templates-io) '
    'and served via a Lambda@URL or CloudFront distribution.'
)

doc.add_heading('7.1 Architecture & Tech Stack', level=2)
add_bullet('HTML5, CSS3, Vanilla JavaScript — no framework dependencies')
add_bullet('Multi-tenant via URL query parameter: ?shop=nissos')
add_bullet('Menu data fetched dynamically from the GetMenuHandler Lambda (GET /menu?shopId=...)')
add_bullet('Cart management via browser Local Storage')
add_bullet('Order submission via POST to CreateOrderHandler (/orders)')

doc.add_heading('7.2 Template Serving via Node.js Lambda', level=2)
doc.add_paragraph(
    'The Node.js Lambda (index.mjs in menu-backend-CICD-Automation) serves the template HTML from S3. '
    'When a customer navigates to the shop URL, the Lambda fetches the shop\'s template.html from S3, '
    'injects shop data as a window.SHOP_DATA JavaScript global, and returns the HTML. '
    'Static assets (CSS, images) are also served through the same Lambda with correct MIME types.'
)

doc.add_heading('7.3 Feature Flags — START ORDER Button', level=2)
doc.add_paragraph(
    'The template checks if the shop has the customerOrdering feature enabled (data.features.customerOrdering === true). '
    'If so, a "START ORDER" button is dynamically injected into the menu page UI. '
    'This button is the entry point to the ordering flow — it redirects the customer to app.html '
    '(in qr-menu-orders-environments), a "chameleon" HTML page that inherits UI elements (theme colors, branding) '
    'from the currently running menu. The shopId is passed as a query parameter so app.html knows which shop\'s menu '
    'and theme to load.'
)

# ═══════════════════════════════════════════════════
# SECTION 8: Ordering & Kitchen Display
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('8. Ordering & Kitchen Display Web Apps', level=1)
add_note('Repository: qr-menu-orders-environments')

doc.add_heading('8.1 Customer Ordering App (app.html) — The Chameleon', level=2)
doc.add_paragraph(
    'A self-contained, single-page ordering application that acts as a "chameleon" — it dynamically inherits '
    'UI elements (theme colors, branding, light/dark mode) from the shop\'s menu that the customer was browsing. '
    'When the customer presses "START ORDER" on the menu page, they are redirected here with the shopId as a query parameter. '
    'app.html fetches the shop\'s menu from the Java GetMenuHandler Lambda (GET /menu?shopId=...), applies the shop\'s '
    'theme settings (primaryColor, themeMode), and renders a clean, animated product listing. '
    'The customer enters their table number, adds items to cart, and submits the order as a POST to /orders on the API Gateway. '
    'Features smooth slide-up animations and a slide-up cart overlay panel.'
)

doc.add_paragraph('Key code reference — Order submission payload:')
add_code('''const payload = {
    shopId: shopId,
    tableNumber: table,
    orderItem: orderItems,    // Array of {name, price, quantity}
    totalAmount: parseFloat(totalAmount.toFixed(2))
    // orderId, createdAt, status="NEW" are generated by Java backend
};

fetch('https://...execute-api.../Prod/orders', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
});''')

doc.add_heading('8.2 Staff Dashboard (staff.html)', level=2)
doc.add_paragraph(
    'This is the staff\'s primary interface — a real-time web-based dashboard designed for display screens at '
    'preparation stations (Bar, Kitchen). Staff do NOT use the React Native mobile app; they use this browser-based dashboard. '
    'The dashboard uses a dark theme (CSS variables: --bg-color: #121212, --surface-color: #1E1E1E, --accent-color: #C4A474) '
    'that matches the MenuAdmin React Native app\'s MD3DarkTheme for visual consistency across the platform. '
    'Uses the Manrope font family for a modern, clean look.'
)
doc.add_paragraph(
    'Each station opens staff.html and sees only the orders relevant to their post. '
    'The dashboard polls GetOrdersHandler (GET /orders?shopId=...) every 10 seconds and renders orders as cards '
    'in a responsive CSS Grid (auto-fill, minmax 300px). Each card shows the table number, timestamp, and ordered items '
    'with quantities. Orders are displayed individually — they do NOT merge even if they are for the same table. '
    'CLAIM notifications are shown only to the waiter who claimed the order. '
    'Staff press "MARK AS DONE" to send a PATCH request to UpdateOrderStatusHandler with the appropriate action payload.'
)

doc.add_paragraph('Key code reference — Polling and action-based status update:')
add_code('''// Fetch orders every 10 seconds
setInterval(fetchOrders, 10000);

// Action-based PATCH (e.g., ITEM_DONE for a specific item)
const payload = { shopId, orderId, action: "ITEM_DONE", productId: "beer_01", itemTimestamp: "..." };
fetch(`${API_BASE_URL}/orders/status`, {
    method: 'PATCH',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify(payload)
});''')

# ═══════════════════════════════════════════════════
# SECTION 9: Admin Mobile App
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('9. Business Owner Mobile Application (MenuAdmin)', level=1)
add_note('Notion Source: 06_Mobile_App_Business_Owner — Repository: MenuAdmin')

doc.add_heading('9.1 Tech Stack', level=2)
add_bullet('React Native with Expo framework')
add_bullet('TypeScript/JavaScript')
add_bullet('React Navigation (Stack Navigator) for screen routing')
add_bullet('React Native Paper (MD3DarkTheme) for Material Design UI components')
add_bullet('Axios for HTTP API communication')

doc.add_heading('9.2 Application Entry Point — App.js', level=2)
add_note('File: MenuAdmin/App.js')
doc.add_paragraph(
    'The app bootstraps with a custom animated splash screen ("MENU ADMINISTRATOR" with "POWERED BY THERISTIS" footer). '
    'It uses Expo SplashScreen to control native splash timing, then runs a 3-second display followed by a 1-second '
    'fade-out animation before revealing the navigation stack. The stack has two screens: Login and MenuDashboard.'
)

doc.add_heading('9.3 LoginScreen.js', level=2)
add_note('File: MenuAdmin/src/screens/LoginScreen.js')
doc.add_paragraph(
    'Presents a clean login form with Shop ID and Password fields. On submit, calls api.login(shopId, password) '
    'which sends a POST to the Node.js Lambda (/login endpoint). On success, navigates to MenuDashboard passing '
    'shopId and password as route parameters.'
)

doc.add_heading('9.4 MenuDashboard.js — Core UI', level=2)
add_note('File: MenuAdmin/src/screens/MenuDashboard.js')
doc.add_paragraph(
    'The main screen of the business owner app. Currently implements menu management functionality '
    '(CRUD for categories and products). '
    'It operates with three view modes: LIST (default accordion view of categories/products), EDIT_CAT (edit/create category form), '
    'and EDIT_PROD (edit/create product form). Key behaviors:'
)
add_bullet('Loads menu data via api.getMenuData(shopId, password) on mount')
add_bullet('Displays categories as expandable Card components with Edit/Add/Delete action buttons')
add_bullet('Products shown inside expanded categories with name, description, and price')
add_bullet('All mutations (save/delete) call api.saveMenuData() which syncs the entire menu JSON back to DynamoDB')
add_bullet('Pull-to-refresh support and hardware back button handling')

doc.add_heading('9.5 Planned Evolution — Business Portfolio', level=2)
doc.add_paragraph(
    'The app is planned to evolve from a simple menu editor into a comprehensive business management portfolio. '
    'Future screens and features include:'
)
add_bullet('Order analytics dashboard: statistics, peak hours, average order value')
add_bullet('Top products report: best sellers by category')
add_bullet('Income & expenses tracking: revenue, costs, profit margins')
add_bullet('Inventory / storage (kava) management: stock tracking for restaurant, cafe, and bar with low-stock alerts')
add_bullet('General business organization tools for the owner\'s operational needs')

doc.add_heading('9.6 API Service Layer — api.js', level=2)
add_note('File: MenuAdmin/src/services/api.js')
doc.add_paragraph(
    'Axios-based wrapper that communicates with the Node.js Lambda (menu-backend-CICD-Automation). '
    'Currently implements three methods:'
)
add_bullet('login(shop_id, password): POST /login — Returns {success: true/false}', 'api.login(): ')
add_bullet('getMenuData(shop_id, password): POST /get-full-data via Lambda Function URL — Returns full shop JSON (menu, settings)', 'api.getMenuData(): ')
add_bullet('saveMenuData(shop_id, password, newData): POST /save-menu via Lambda Function URL — Persists updated menu to DynamoDB', 'api.saveMenuData(): ')

doc.add_paragraph(
    'In the future, this service may be extended with calls to the Java API Gateway endpoints for analytics and reporting '
    '(e.g., fetching order history, revenue data). Staff order management is handled via the web-based staff.html dashboard, '
    'not through this mobile app.'
)

# ═══════════════════════════════════════════════════
# SECTION 10: Terraform
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('10. Infrastructure as Code (Terraform)', level=1)
add_note('Notion Source: 02_Cloud_Architecture_And_DevOps — Repository: menu-backend-CICD-Automation')

doc.add_heading('10.1 Terraform Configuration — main.tf', level=2)
add_note('File: menu-backend-CICD-Automation/terraform/main.tf')
doc.add_paragraph(
    'Manages the Node.js Lambda function and its infrastructure. Key design decisions:'
)

add_bullet('Backend state stored in S3 (bucket: qr-menu-terraform-state-2026, key: qr-menu/terraform.tfstate)')
add_bullet('Uses a variable env_stage (values: "dev" or "prod") to control deployment behavior')
add_bullet('Lambda function: "qr-menu", Node.js 20.x runtime, referencing an existing IAM role')
add_bullet('The source code is zipped automatically via the archive_file data source, excluding .git, terraform, etc.')
add_bullet('publish flag is conditional: only creates numbered versions in prod (publish = var.env_stage == "prod" ? true : false)')
add_bullet('A "PROD" alias resource points to the latest published version, protecting production from dev changes')

doc.add_paragraph('Key code reference — Conditional versioning:')
add_code('''resource "aws_lambda_function" "qr_menu" {
  function_name = "qr-menu"
  runtime       = "nodejs20.x"
  handler       = "index.handler"
  # Only publish a numbered version in prod
  publish = var.env_stage == "prod" ? true : false
  environment {
    variables = {
      TABLE_NAME = "menus"
      STAGE      = var.env_stage
    }
  }
}

resource "aws_lambda_alias" "prod_alias" {
  name             = "PROD"
  function_name    = aws_lambda_function.qr_menu.function_name
  function_version = aws_lambda_function.qr_menu.version
}''')

doc.add_heading('10.2 Lambda Runtime Source Code', level=2)
doc.add_paragraph(
    'These files are NOT provisioning or configuration scripts. They are the runtime source code '
    'of the Node.js Lambda function ("qr-menu") and are deployed as the Lambda\'s code package via Terraform.'
)

doc.add_heading('10.2.1 db.mjs — Data Access Layer', level=3)
add_note('File: menu-backend-CICD-Automation/db.mjs')
doc.add_paragraph(
    'Provides a clean Data Access Layer using AWS SDK v3 (DynamoDBDocumentClient). '
    'Exports two functions: getShopData(shopId) for reading shop configurations, and '
    'saveShopData(shopData) for writing/updating. Uses the "Menus" table in eu-central-1.'
)

doc.add_heading('10.2.2 index.mjs — Lambda Handler & Router', level=3)
add_note('File: menu-backend-CICD-Automation/index.mjs')
doc.add_paragraph(
    'The main entry point for the Node.js Lambda. Implements a custom router that handles:'
)
add_bullet('GET /get-menu?shopId=X: Public endpoint. Fetches menu and settings from DynamoDB.')
add_bullet('POST /login: Validates shop_id and password against DynamoDB records.')
add_bullet('POST /save-menu: Protected. Re-verifies password before writing updated menu to DynamoDB.')
add_bullet('POST /get-full-data: Protected. Returns menu + settings after password verification.')
add_bullet('S3 File Serving: For any non-API path, serves static files (HTML, CSS, images) from the S3 bucket. '
           'HTML files get shop data injected as a window.SHOP_DATA script tag before serving.')

# ═══════════════════════════════════════════════════
# SECTION 11: CI/CD
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('11. CI/CD Pipelines (GitHub Actions)', level=1)
add_note('Notion Source: 02_Cloud_Architecture_And_DevOps')

doc.add_heading('11.1 Frontend Deployment — deploy-frontend.yml', level=2)
add_note('File: menu-frontend-CICD-Automation/.github/workflows/deploy-frontend.yml')
doc.add_paragraph('Triggers on push to main branch when files under shops/ change. Steps:')
add_bullet('Checkout code')
add_bullet('Configure AWS credentials from GitHub Secrets')
add_bullet('Sync the shops/ directory to S3 bucket (qr-templates-io) with --delete and no-cache headers')

doc.add_paragraph('Key command:')
add_code('aws s3 sync ./shops s3://qr-templates-io --delete --exclude "terraform/*" --cache-control "max-age=0, no-cache, no-store, must-revalidate"')

doc.add_heading('11.2 Backend Deployment — deploy.yml', level=2)
add_note('File: menu-backend-CICD-Automation/.github/workflows/deploy.yml')
doc.add_paragraph(
    'Triggers on push to both main and dev branches. Implements a multi-environment deployment strategy:'
)
add_bullet('Sets ENV_NAME to "prod" if branch is main, "dev" otherwise')
add_bullet('Runs terraform init and terraform apply with -var="env_stage=$ENV_NAME"')
add_bullet('For prod (main branch): publish=true creates a numbered Lambda version, PROD alias updates to new version')
add_bullet('For dev (dev branch): publish=false uploads code to $LATEST only, PROD alias stays on the old stable version')

doc.add_paragraph('Key code reference — Environment-aware deployment:')
add_code('''# Set environment based on branch
if [ "${{ github.ref_name }}" == "main" ]; then
  echo "ENV_NAME=prod" >> $GITHUB_ENV
else
  echo "ENV_NAME=dev" >> $GITHUB_ENV
fi

# Apply with environment variable
terraform apply -auto-approve -var="env_stage=${{ env.ENV_NAME }}"''')

doc.add_heading('11.3 Java Backend Deployment (SAM)', level=2)
doc.add_paragraph(
    'The Java-based qr-menu-order-service is deployed using AWS SAM CLI. The samconfig.toml file contains '
    'deployment configuration. The workflow is: mvn package builds the JAR -> sam build -> sam deploy pushes '
    'the CloudFormation stack with the Lambda functions and API Gateway to AWS.'
)

# ═══════════════════════════════════════════════════
# SECTION 12: Data Flow
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('12. Data Flow & Component Interaction', level=1)

doc.add_heading('12.1 Customer Order Flow', level=2)
doc.add_paragraph('Step-by-step data flow when a customer places an order:')

steps = [
    'Customer scans QR code -> Opens template.html via Node.js Lambda (S3 serving)',
    'Menu page loads -> Node.js Lambda fetches shop data from DynamoDB "Menus" table and injects it as window.SHOP_DATA',
    'Customer browses the menu. Presses "START ORDER" button embedded in the menu UI',
    'Redirected to app.html (qr-menu-orders-environments S3 bucket) — a chameleon page that inherits UI theme from the running menu',
    'app.html fetches menu via Java GetMenuHandler (GET /menu?shopId=X) and applies the shop\'s theme (colors, light/dark mode)',
    'Customer adds items to cart and enters table number',
    'Customer submits order -> POST /orders hits Java CreateOrderHandler',
    'CreateOrderHandler generates orderId (ORDER#TIMESTAMP#UUID), sets status=NEW, writes to DynamoDB "orders" table',
    'Order now visible on the staff dashboard (staff.html)',
]
for i, step in enumerate(steps, 1):
    add_bullet(f'{step}', f'Step {i}: ')

doc.add_heading('12.2 Staff Order Processing Flow', level=2)
steps2 = [
    'Staff opens staff.html in a browser at their preparation station (Bar or Kitchen)',
    'staff.html polls GetOrdersHandler (GET /orders?shopId=X) every 10 seconds -> Gets all orders from DynamoDB',
    'Each station sees only orders filtered by their station field (BAR/KITCHEN)',
    'Station staff marks items done -> PATCH /orders/status transitions order status',
    'When all items across all stations are DONE -> parent status becomes READY',
    'Waiter delivers the order and closes it after payment -> status=CLOSED, paymentStatus=PAID',
]
for i, step in enumerate(steps2, 1):
    add_bullet(f'{step}', f'Step {i}: ')

doc.add_heading('12.3 Menu Management Flow', level=2)
steps3 = [
    'Business owner opens MenuAdmin app on their phone -> Logs in with Shop ID and Password',
    'MenuDashboard calls api.getMenuData() -> POST /get-full-data to Node.js Lambda via Function URL',
    'Node.js Lambda reads from DynamoDB "Menus" table and returns menu JSON',
    'Owner adds/edits/deletes categories or products in the UI',
    'On save, api.saveMenuData() -> POST /save-menu to Node.js Lambda via Function URL',
    'Node.js Lambda writes the complete updated JSON back to DynamoDB (PutCommand)',
    'Customer frontend and ordering app (app.html) automatically serve updated data on next page load',
]
for i, step in enumerate(steps3, 1):
    add_bullet(f'{step}', f'Step {i}: ')

doc.add_heading('12.4 Two Lambda Systems — How They Coexist', level=2)
doc.add_paragraph(
    'The platform uses two distinct Lambda systems that share the same DynamoDB "Menus" table but serve different purposes:'
)

coexist_table = doc.add_table(rows=3, cols=4)
coexist_table.style = 'Light Shading Accent 1'
for i, h in enumerate(['Aspect', 'Node.js Lambda (qr-menu)', 'Java Lambdas (SAM)', 'Shared Resource']):
    coexist_table.rows[0].cells[i].text = h
    for run in coexist_table.rows[0].cells[i].paragraphs[0].runs:
        run.bold = True

coexist_data = [
    ['Purpose', 'Menu management, owner auth, S3 template serving', 'Order processing (CRUD), menu reading for customers', 'DynamoDB "Menus" table'],
    ['Consumers', 'Business owner mobile app (MenuAdmin), customer template pages', 'Customer ordering app (app.html), staff dashboard (staff.html)', 'DynamoDB "orders" table (Java only)'],
]
for row_idx, row_data in enumerate(coexist_data, 1):
    for col_idx, val in enumerate(row_data):
        coexist_table.rows[row_idx].cells[col_idx].text = val

# ═══════════════════════════════════════════════════
# SECTION 13: Epics & Roadmap
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('13. Epic & Kanban Roadmap', level=1)
add_note('Notion Source: 00_Master_Business_Plan — Section 4')

doc.add_heading('Epic 1: Database & Order Lifecycle Schema', level=2)
add_bullet('Done', 'Status: ')
add_bullet('Objective: Migrate DynamoDB fields to support claimedBy, station, paymentStatus, and itemStatus.')
add_bullet('Impact: Enables the full order lifecycle (NEW -> CLAIMED -> PARTIAL -> READY -> CLOSED) and station-based routing.')
add_bullet('Implementation: Order.java extended with claimedBy, paymentStatus. OrderItem.java extended with station, '
           'itemStatus, timestamp. CreateOrderHandler sets defaults on creation: status="NEW", paymentStatus="UNPAID", '
           'claimedBy="NONE", itemStatus="PENDING" per item, station defaults to "KITCHEN" if null.')

doc.add_heading('Epic 2: Smart Backend Routing (Java)', level=2)
add_bullet('Done', 'Status: ')
add_bullet('Objective: Refactor UpdateOrderStatusHandler to support action-based payloads (CLAIM, ITEM_DONE, APPEND, CLOSE).')
add_bullet('Impact: Transforms the backend from a simple status updater to an intelligent order routing engine.')
add_bullet('Implementation: UpdateOrderStatusHandler.java fully refactored with action-based switch routing. '
           'New OrderUpdateRequest.java model handles all action payloads. Uses read-modify-write pattern '
           '(GetItem -> modify items array -> PutItem). ITEM_DONE uses productId+itemTimestamp dual key for disambiguation. '
           'SAM template updated: DynamoDBCrudPolicy for GetItem+PutItem access, global CORS configuration.')

doc.add_heading('Epic 3: Business Owner App — Portfolio Evolution', level=2)
add_bullet('Status: To Do', 'Status: ')
add_bullet('Objective: Evolve MenuAdmin from a menu editor into a full business management portfolio for the restaurant owner.')
add_bullet('Impact: Gives the business owner analytics, financial insights, and inventory control in a single mobile app.')
add_bullet('Code Impact: Add new screens to MenuAdmin for order analytics, top products, income/expenses, '
           'and inventory management. Extend api.js with new methods for fetching order history and reporting data. '
           'Implement data visualization components.')

doc.add_heading('Epic 4: Dynamic Frontend Templating', level=2)
add_bullet('Status: To Do', 'Status: ')
add_bullet('Objective: Refactor menu-frontend-CICD-Automation to prevent HTML duplication across shops.')
add_bullet('Impact: Reduces maintenance overhead. New shops can be onboarded by adding a config file instead of duplicating HTML.')
add_bullet('Code Impact: Create a base template and a build script (in the GitHub Action CI pipeline) that injects '
           'shop-specific metadata (prices, items, logos, themes) during deployment.')

# ═══════════════════════════════════════════════════
# FOOTER
# ═══════════════════════════════════════════════════
doc.add_page_break()
doc.add_heading('Appendix: Repository Structure', level=1)

repos = {
    'qr-menu-order-service (Java Backend)': [
        'src/main/java/com/qrmenu/orders/handlers/CreateOrderHandler.java',
        'src/main/java/com/qrmenu/orders/handlers/GetOrdersHandler.java',
        'src/main/java/com/qrmenu/orders/handlers/UpdateOrderStatusHandler.java',
        'src/main/java/com/qrmenu/orders/handlers/GetMenuHandler.java',
        'src/main/java/com/qrmenu/orders/models/Order.java',
        'src/main/java/com/qrmenu/orders/models/OrderItem.java',
        'src/main/java/com/qrmenu/orders/models/OrderUpdateRequest.java',
        'template.yaml (SAM infrastructure)',
        'pom.xml (Maven dependencies)',
    ],
    'menu-backend-CICD-Automation (Node.js + Terraform)': [
        'index.mjs (Lambda handler & router)',
        'db.mjs (DynamoDB data access layer)',
        'terraform/main.tf (IaC configuration)',
        '.github/workflows/deploy.yml (CI/CD pipeline)',
    ],
    'menu-frontend-CICD-Automation (Customer Web)': [
        'shops/nissos/template.html + style.css',
        'shops/rakoumel/template.html + style.css',
        'shops/theristis/template.html + style.css',
        '.github/workflows/deploy-frontend.yml (S3 sync)',
    ],
    'MenuAdmin (Business Owner Mobile App)': [
        'App.js (Entry point & navigation)',
        'src/screens/LoginScreen.js',
        'src/screens/MenuDashboard.js (Menu CRUD)',
        'src/services/api.js (API wrapper)',
    ],
    'qr-menu-orders-environments (Ordering & Staff Web Apps)': [
        'app.html (Customer ordering app — chameleon UI)',
        'staff.html (Staff dashboard — preparation stations)',
    ],
}

for repo, files in repos.items():
    doc.add_heading(repo, level=2)
    for f in files:
        add_bullet(f)

# ─── Save ───
output_path = '/home/user/QR_Menu_Platform_Documentation.docx'
doc.save(output_path)
print(f'Document saved to: {output_path}')
