import csv
from collections import defaultdict
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.uix.filechooser import FileChooserListView
from kivy.uix.popup import Popup
from kivy.core.window import Window
from plyer import filechooser

# -------------------------------
# CONFIGURATION (same as your script)
# -------------------------------
COST_MAP = {
    'Cheap Stylus': 100,
    'Expensive Stylus': 650,
    '2-in-1 Card Reader': 170,
    '3-in-1 Card Reader': 170,
    'Circuit Breaker': 1800,
    'Smart Plug': 1100,
    'Otg Adapter': 170,
}

def categorize(product_name, price_paid):
    if 'Stylus' in product_name or 'Stylus Pen' in product_name:
        return 'Cheap Stylus' if price_paid < 700 else 'Expensive Stylus'
    elif 'Memory Card Reader' in product_name:
        return '2-in-1 Card Reader' if '2 In1' in product_name else '3-in-1 Card Reader'
    elif 'Circuit Breaker' in product_name:
        return 'Circuit Breaker'
    elif 'Smart Socket' in product_name:
        return 'Smart Plug'
    elif 'Otg Adapter' in product_name or 'OTG' in product_name:
        return 'Otg Adapter'
    else:
        return 'Other'

def process_file(filename):
    orders = {}
    period = "Unknown"
    try:
        with open(filename, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                if period == "Unknown" and 'Statement Period' in row:
                    period = row['Statement Period']
                order_id = row['Order Line ID']
                amount = float(row['Amount(Include Tax)'].replace(',', '').strip())
                fee_name = row['Fee Name']
                product_name = row['Product Name']
                if order_id not in orders:
                    orders[order_id] = {
                        'total': 0.0,
                        'product_name': product_name,
                        'price_paid': 0.0,
                    }
                orders[order_id]['total'] += amount
                if fee_name == 'Product Price Paid by Buyer':
                    orders[order_id]['price_paid'] = amount
                if product_name:
                    orders[order_id]['product_name'] = product_name
    except Exception as e:
        return None, str(e)

    positive_payouts = defaultdict(list)
    negative_payouts = defaultdict(list)
    all_categories = set()

    for data in orders.values():
        net = data['total']
        product = data['product_name']
        price = data['price_paid']
        cat = categorize(product, price)
        all_categories.add(cat)
        if net < 0:
            negative_payouts[cat].append(net)
        else:
            positive_payouts[cat].append(net)

    total_positive_profit = 0.0
    category_positive_profit = {}
    for cat in all_categories:
        pos_list = positive_payouts.get(cat, [])
        count_pos = len(pos_list)
        sum_pos = sum(pos_list)
        cost = COST_MAP.get(cat, 0)
        gross_profit = sum_pos - (count_pos * cost)
        category_positive_profit[cat] = gross_profit
        total_positive_profit += gross_profit

    total_negative = sum(sum(lst) for lst in negative_payouts.values())
    net_profit = total_positive_profit + total_negative

    return {
        'period': period,
        'positive_payouts': positive_payouts,
        'negative_payouts': negative_payouts,
        'category_positive_profit': category_positive_profit,
        'total_positive_profit': total_positive_profit,
        'total_negative': total_negative,
        'net_profit': net_profit,
    }, None

# -------------------------------
# Kivy App
# -------------------------------
class ProfitApp(App):
    def build(self):
        self.files = []
        self.output_text = ""

        # Main layout
        main_layout = BoxLayout(orientation='vertical', padding=10, spacing=10)
        main_layout.add_widget(Label(text='📊 Profit Calculator', font_size='24sp', size_hint_y=None, height=50))

        # File selection button
        select_btn = Button(text='Select CSV Files', size_hint_y=None, height=50)
        select_btn.bind(on_press=self.select_files)
        main_layout.add_widget(select_btn)

        # File list display (scrollable)
        self.file_list_label = Label(text='No files selected', size_hint_y=None, height=200, halign='left', valign='top')
        self.file_list_label.bind(size=self.file_list_label.setter('text_size'))
        file_scroll = ScrollView(size_hint_y=None, height=200)
        file_scroll.add_widget(self.file_list_label)
        main_layout.add_widget(file_scroll)

        # Calculate button
        calc_btn = Button(text='Calculate Profit', size_hint_y=None, height=50)
        calc_btn.bind(on_press=self.calculate)
        main_layout.add_widget(calc_btn)

        # Output area (scrollable)
        self.output = TextInput(text='', readonly=True, multiline=True, font_size='14sp')
        output_scroll = ScrollView()
        output_scroll.add_widget(self.output)
        main_layout.add_widget(output_scroll)

        return main_layout

    def select_files(self, instance):
        # Use plyer filechooser for multiple file selection
        filechooser.open_file(multiple=True, on_selection=self.file_selected)

    def file_selected(self, selection):
        if selection:
            self.files = selection
            # Update label
            names = '\n'.join([f'• {f.split("/")[-1]}' for f in self.files])
            self.file_list_label.text = f"Selected files:\n{names}"
        else:
            self.file_list_label.text = "No files selected"

    def calculate(self, instance):
        if not self.files:
            self.output.text = "Please select at least one CSV file."
            return

        full_output = ""
        grand_total_net = 0.0
        all_results = []

        for fname in self.files:
            result, error = process_file(fname)
            if error:
                full_output += f"Error processing {fname}: {error}\n\n"
                continue

            period = result['period']
            pos = result['positive_payouts']
            neg = result['negative_payouts']
            cat_profit = result['category_positive_profit']
            total_pos_profit = result['total_positive_profit']
            total_neg = result['total_negative']
            net = result['net_profit']

            full_output += f"📁 File: {fname.split('/')[-1]}\n"
            full_output += f"   Period: {period}\n"
            full_output += "   Positive payouts per category:\n"
            for cat in sorted(pos.keys()):
                if pos[cat]:
                    full_output += f"      {cat}: {len(pos[cat])} orders, total {sum(pos[cat]):.2f}\n"
            full_output += "   Negative payouts per category:\n"
            for cat in sorted(neg.keys()):
                if neg[cat]:
                    full_output += f"      {cat}: {len(neg[cat])} orders, total {sum(neg[cat]):.2f}\n"
            full_output += "   Gross profit from positive orders:\n"
            for cat, profit in sorted(cat_profit.items()):
                full_output += f"      {cat}: {profit:.2f}\n"
            full_output += f"   Total gross profit: {total_pos_profit:.2f}\n"
            full_output += f"   Total negative payouts: {total_neg:.2f}\n"
            full_output += f"   Net profit: {net:.2f}\n\n"

            all_results.append((period, net))
            grand_total_net += net

        # Combined results
        full_output += "="*50 + "\n"
        full_output += "📊 COMBINED RESULTS\n"
        for period, profit in all_results:
            full_output += f"   {period}: {profit:.2f}\n"
        full_output += f"\nGrand Total Net Profit (before 5% deduction): {grand_total_net:.2f}\n"
        loss = grand_total_net * 0.05
        final_profit = grand_total_net - loss
        full_output += f"Less 5% loss: -{loss:.2f}\n"
        full_output += f"FINAL NET PROFIT: {final_profit:.2f}\n"

        self.output.text = full_output

if __name__ == '__main__':
    ProfitApp().run()