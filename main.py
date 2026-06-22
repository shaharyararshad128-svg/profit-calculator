# main.py – Kivy Profit Calculator (Android-ready) with Product Management

import csv
import json
import os
from collections import defaultdict
from kivy.app import App
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.button import Button
from kivy.uix.label import Label
from kivy.uix.scrollview import ScrollView
from kivy.uix.textinput import TextInput
from kivy.uix.screenmanager import ScreenManager, Screen
from kivy.uix.gridlayout import GridLayout
from kivy.uix.widget import Widget
from kivy.graphics import Color, RoundedRectangle
from kivy.metrics import dp
from kivy.core.window import Window
from plyer import filechooser

# -------------------------------
# Default cost map (fallback)
# -------------------------------
DEFAULT_COST_MAP = {
    'Cheap Stylus': 100,
    'Expensive Stylus': 650,
    '2-in-1 Card Reader': 170,
    '3-in-1 Card Reader': 170,
    'Circuit Breaker': 1800,
    'Smart Plug': 1100,
    'Otg Adapter': 170,
}

# -------------------------------
# Product Manager (load/save cost map)
# -------------------------------
class ProductManager:
    def __init__(self, app):
        self.app = app
        self.cost_map = {}
        self.load()

    def get_data_file(self):
        return os.path.join(self.app.user_data_dir, 'product_costs.json')

    def load(self):
        try:
            with open(self.get_data_file(), 'r') as f:
                self.cost_map = json.load(f)
        except (FileNotFoundError, json.JSONDecodeError):
            self.cost_map = DEFAULT_COST_MAP.copy()
            self.save()

    def save(self):
        with open(self.get_data_file(), 'w') as f:
            json.dump(self.cost_map, f, indent=2)

    def get_cost(self, category):
        return self.cost_map.get(category, 0)

    def set_cost(self, category, cost):
        self.cost_map[category] = cost
        self.save()

    def add_product(self, category, cost):
        self.cost_map[category] = cost
        self.save()

    def remove_product(self, category):
        if category in self.cost_map:
            del self.cost_map[category]
            self.save()

    def get_all_categories(self):
        return sorted(self.cost_map.keys())

# -------------------------------
# CSV Processing (uses current cost map)
# -------------------------------
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

def process_file(filename, cost_map):
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
                    orders[order_id] = {'total': 0.0, 'product_name': product_name, 'price_paid': 0.0}
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
        cost = cost_map.get(cat, 0)
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
# Custom Widgets for Results
# -------------------------------
class CategoryCard(BoxLayout):
    def __init__(self, cat, pos_list, neg_list, gross_profit, **kwargs):
        super().__init__(**kwargs)
        self.orientation = 'vertical'
        self.size_hint_y = None
        self.height = dp(200)  # auto-adjust later
        self.padding = dp(10)
        self.spacing = dp(5)

        # Background (rounded rectangle)
        with self.canvas.before:
            Color(0.95, 0.95, 0.95, 1)
            self.rect = RoundedRectangle(pos=self.pos, size=self.size, radius=[dp(10)])
        self.bind(pos=self.update_rect, size=self.update_rect)

        # Category title
        title = Label(text=f"[b]{cat}[/b]", markup=True, font_size='18sp', size_hint_y=None, height=dp(30))
        self.add_widget(title)

        # Details grid
        grid = GridLayout(cols=2, size_hint_y=None, height=dp(120), spacing=dp(5))
        grid.add_widget(Label(text="Positive orders:", font_size='14sp', halign='left'))
        grid.add_widget(Label(text=f"{len(pos_list)}  (total {sum(pos_list):.2f})", font_size='14sp', halign='left'))
        grid.add_widget(Label(text="Negative orders:", font_size='14sp', halign='left'))
        grid.add_widget(Label(text=f"{len(neg_list)}  (total {sum(neg_list):.2f})", font_size='14sp', halign='left'))
        grid.add_widget(Label(text="Gross profit:", font_size='14sp', halign='left'))
        grid.add_widget(Label(text=f"{gross_profit:.2f}", font_size='14sp', color=(0,0.6,0,1) if gross_profit>=0 else (0.8,0,0,1)))
        self.add_widget(grid)

        # Adjust height based on content
        self.height = dp(30) + dp(120) + dp(20)

    def update_rect(self, *args):
        self.rect.pos = self.pos
        self.rect.size = self.size

# -------------------------------
# Screens
# -------------------------------
class MainScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.files = []
        self.build_ui()

    def build_ui(self):
        layout = BoxLayout(orientation='vertical', padding=dp(20), spacing=dp(15))
        layout.add_widget(Label(text='📊 Profit Calculator', font_size='28sp', size_hint_y=None, height=dp(60), color=(0.2,0.2,0.8,1)))

        # File selection
        select_btn = Button(text='Select CSV Files', size_hint_y=None, height=dp(50), background_color=(0.1,0.5,0.8,1))
        select_btn.bind(on_press=self.select_files)
        layout.add_widget(select_btn)

        self.file_list_label = Label(text='No files selected', size_hint_y=None, height=dp(150), halign='left', valign='top')
        self.file_list_label.bind(size=self.file_list_label.setter('text_size'))
        file_scroll = ScrollView(size_hint_y=None, height=dp(150))
        file_scroll.add_widget(self.file_list_label)
        layout.add_widget(file_scroll)

        # Calculate button
        calc_btn = Button(text='Calculate Profit', size_hint_y=None, height=dp(50), background_color=(0.2,0.7,0.3,1))
        calc_btn.bind(on_press=self.calculate)
        layout.add_widget(calc_btn)

        # Manage products button
        manage_btn = Button(text='Manage Products', size_hint_y=None, height=dp(50), background_color=(0.6,0.4,0.8,1))
        manage_btn.bind(on_press=self.go_to_manage)
        layout.add_widget(manage_btn)

        self.add_widget(layout)

    def select_files(self, instance):
        filechooser.open_file(multiple=True, on_selection=self.file_selected)

    def file_selected(self, selection):
        if selection:
            self.files = selection
            names = '\n'.join([f'• {f.split("/")[-1]}' for f in self.files])
            self.file_list_label.text = f"Selected files:\n{names}"
        else:
            self.file_list_label.text = "No files selected"

    def calculate(self, instance):
        if not self.files:
            # show a simple popup or message (we'll just go to results with error)
            results_screen = self.manager.get_screen('results')
            results_screen.set_error("Please select at least one CSV file.")
            self.manager.current = 'results'
            return

        app = App.get_running_app()
        cost_map = app.product_manager.cost_map

        full_results = []
        grand_total_net = 0.0
        errors = []

        for fname in self.files:
            result, error = process_file(fname, cost_map)
            if error:
                errors.append(f"{fname}: {error}")
                continue
            full_results.append((fname, result))
            grand_total_net += result['net_profit']

        # Pass to results screen
        results_screen = self.manager.get_screen('results')
        if errors:
            results_screen.set_error("\n".join(errors))
        else:
            results_screen.display_results(full_results, grand_total_net)
        self.manager.current = 'results'

    def go_to_manage(self, instance):
        self.manager.current = 'manage'

class ResultsScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.build_ui()

    def build_ui(self):
        layout = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))
        # Header
        header = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        header.add_widget(Label(text='📈 Results', font_size='24sp'))
        back_btn = Button(text='Back', size_hint_x=0.2, background_color=(0.6,0.6,0.6,1))
        back_btn.bind(on_press=self.go_back)
        header.add_widget(back_btn)
        layout.add_widget(header)

        # Scrollable content
        self.scroll = ScrollView()
        self.content = BoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(10), padding=dp(10))
        self.content.bind(minimum_height=self.content.setter('height'))
        self.scroll.add_widget(self.content)
        layout.add_widget(self.scroll)

        self.add_widget(layout)

    def set_error(self, message):
        self.content.clear_widgets()
        self.content.add_widget(Label(text=f"❌ Error:\n{message}", font_size='16sp', color=(0.8,0,0,1)))

    def display_results(self, full_results, grand_total_net):
        self.content.clear_widgets()

        for fname, result in full_results:
            # File header
            file_label = Label(text=f"📁 {fname.split('/')[-1]}   Period: {result['period']}", font_size='18sp', bold=True, size_hint_y=None, height=dp(40))
            self.content.add_widget(file_label)

            # Category cards
            cat_profit = result['category_positive_profit']
            pos = result['positive_payouts']
            neg = result['negative_payouts']
            all_cats = sorted(set(list(cat_profit.keys()) + list(pos.keys()) + list(neg.keys())))
            for cat in all_cats:
                pos_list = pos.get(cat, [])
                neg_list = neg.get(cat, [])
                gross = cat_profit.get(cat, 0.0)
                card = CategoryCard(cat, pos_list, neg_list, gross)
                self.content.add_widget(card)

            # Summary for this file
            summary = f"Total gross profit: {result['total_positive_profit']:.2f}   |   Total negative: {result['total_negative']:.2f}   |   Net profit: {result['net_profit']:.2f}"
            self.content.add_widget(Label(text=summary, font_size='16sp', size_hint_y=None, height=dp(30), color=(0.1,0.1,0.7,1)))
            self.content.add_widget(Widget(size_hint_y=None, height=dp(10)))  # spacer

        # Grand total summary
        loss = grand_total_net * 0.05
        final_profit = grand_total_net - loss
        grand_text = f"Grand Total Net Profit (before deduction): {grand_total_net:.2f}\nLess 5% loss: -{loss:.2f}\nFINAL NET PROFIT: {final_profit:.2f}"
        self.content.add_widget(Label(text=grand_text, font_size='20sp', bold=True, size_hint_y=None, height=dp(100), color=(0,0.5,0,1)))

    def go_back(self, instance):
        self.manager.current = 'main'

class ManageProductsScreen(Screen):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.build_ui()

    def build_ui(self):
        layout = BoxLayout(orientation='vertical', padding=dp(10), spacing=dp(10))
        # Header
        header = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        header.add_widget(Label(text='⚙️ Manage Products', font_size='24sp'))
        back_btn = Button(text='Back', size_hint_x=0.2, background_color=(0.6,0.6,0.6,1))
        back_btn.bind(on_press=self.go_back)
        header.add_widget(back_btn)
        layout.add_widget(header)

        # List of products with cost and remove
        self.product_list = BoxLayout(orientation='vertical', size_hint_y=None, spacing=dp(5))
        self.product_list.bind(minimum_height=self.product_list.setter('height'))
        scroll = ScrollView()
        scroll.add_widget(self.product_list)
        layout.add_widget(scroll)

        # Add new product
        add_box = BoxLayout(size_hint_y=None, height=dp(50), spacing=dp(10))
        self.new_name = TextInput(hint_text='Product name', multiline=False)
        self.new_cost = TextInput(hint_text='Cost price', multiline=False, input_filter='float')
        add_btn = Button(text='Add', size_hint_x=0.3)
        add_btn.bind(on_press=self.add_product)
        add_box.add_widget(self.new_name)
        add_box.add_widget(self.new_cost)
        add_box.add_widget(add_btn)
        layout.add_widget(add_box)

        self.add_widget(layout)
        self.populate_products()

    def on_enter(self):
        self.populate_products()

    def populate_products(self):
        self.product_list.clear_widgets()
        app = App.get_running_app()
        cost_map = app.product_manager.cost_map
        for cat, cost in sorted(cost_map.items()):
            row = BoxLayout(size_hint_y=None, height=dp(40), spacing=dp(5))
            label = Label(text=cat, size_hint_x=0.4)
            cost_input = TextInput(text=str(cost), multiline=False, input_filter='float', size_hint_x=0.3)
            cost_input.bind(on_text_validate=self.update_cost)  # update on enter
            cost_input.bind(focus=self.on_cost_focus)  # also on losing focus
            cost_input.product_category = cat
            remove_btn = Button(text='X', size_hint_x=0.15, background_color=(0.8,0.2,0.2,1))
            remove_btn.bind(on_press=self.remove_product)
            remove_btn.product_category = cat
            row.add_widget(label)
            row.add_widget(cost_input)
            row.add_widget(remove_btn)
            self.product_list.add_widget(row)

    def update_cost(self, instance):
        # called when user presses enter
        self.save_cost(instance)

    def on_cost_focus(self, instance, value):
        if not value:  # lost focus
            self.save_cost(instance)

    def save_cost(self, instance):
        try:
            new_cost = float(instance.text)
        except ValueError:
            return
        cat = instance.product_category
        app = App.get_running_app()
        app.product_manager.set_cost(cat, new_cost)

    def remove_product(self, instance):
        cat = instance.product_category
        app = App.get_running_app()
        app.product_manager.remove_product(cat)
        self.populate_products()

    def add_product(self, instance):
        name = self.new_name.text.strip()
        cost_text = self.new_cost.text.strip()
        if not name or not cost_text:
            return
        try:
            cost = float(cost_text)
        except ValueError:
            return
        app = App.get_running_app()
        app.product_manager.add_product(name, cost)
        self.new_name.text = ''
        self.new_cost.text = ''
        self.populate_products()

    def go_back(self, instance):
        self.manager.current = 'main'

# -------------------------------
# The App
# -------------------------------
class ProfitApp(App):
    def build(self):
        # Initialize product manager
        self.product_manager = ProductManager(self)

        sm = ScreenManager()
        sm.add_widget(MainScreen(name='main'))
        sm.add_widget(ResultsScreen(name='results'))
        sm.add_widget(ManageProductsScreen(name='manage'))
        return sm

if __name__ == '__main__':
    ProfitApp().run()