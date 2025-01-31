<?php
/*
Plugin Name: Geniki Taxydromiki Voucher Export
Plugin URI: 
Description: Export voucher numbers based on creation date for Geniki Taxydromiki
Version: 1.0
Author: 
Text Domain: geniki-voucher-export
Requires PHP: 7.4
*/

// Security check
defined('ABSPATH') || exit;

// Load Composer dependencies
require_once __DIR__ . '/vendor/autoload.php';

// Plugin class
class Geniki_Voucher_Export {

    // Initialize plugin
    public function __construct() {
        add_action('admin_menu', [$this, 'add_admin_menu']);
    }

    // Add admin menu
    public function add_admin_menu() {
        add_submenu_page(
            'woocommerce',
            __('Voucher Export', 'geniki-voucher-export'),
            __('Voucher Export', 'geniki-voucher-export'),
            'manage_woocommerce',
            'geniki-voucher-export',
            [$this, 'render_export_page']
        );
    }

    // Render export page
    public function render_export_page() {
        ?>
        <div class="wrap">
            <h1><?php esc_html_e('Export Vouchers', 'geniki-voucher-export');?></h1>
            <form method="post">
                <?php wp_nonce_field('geniki_export_nonce', 'geniki_nonce'); ?>
                <label for="start_date"><?php esc_html_e('Start Date:', 'geniki-voucher-export'); ?></label>
                <input type="date" name="start_date" required>
                
                <label for="end_date"><?php esc_html_e('End Date:', 'geniki-voucher-export'); ?></label>
                <input type="date" name="end_date" required>
                
                <input type="submit" name="export_vouchers" 
                       class="button button-primary" 
                       value="<?php esc_attr_e('Export', 'geniki-voucher-export'); ?>">
            </form>
        </div>
        <?php
    }

    // Process export
    public function process_export() {
        if (!isset($_POST['export_vouchers'])) return;

        // Security checks
        check_admin_referer('geniki_export_nonce', 'geniki_nonce');
        if (!current_user_can('manage_woocommerce')) wp_die(__('Permission denied', 'geniki-voucher-export'));

        // Get dates
        $start_date = sanitize_text_field($_POST['start_date'] . ' 00:00:00');
        $end_date = sanitize_text_field($_POST['end_date'] . ' 23:59:59');

        // Get voucher data
        $data = $this->get_voucher_data($start_date, $end_date);
        
        if (empty($data)) {
            wp_die(__('No vouchers found in selected date range', 'geniki-voucher-export'));
        }

        // Generate XLSX
        $this->generate_xlsx($data);
    }

    // Get voucher data
    private function get_voucher_data($start_date, $end_date) {
        global $wpdb;

        $query = $wpdb->prepare(
            "SELECT c.comment_post_ID as order_id, 
                    c.comment_date as voucher_date,
                    c.comment_content as note_content
             FROM {$wpdb->comments} c
             WHERE c.comment_type = 'order_note'
             AND c.comment_content LIKE '%Geniki%'
             AND c.comment_date >= %s
             AND c.comment_date <= %s
             ORDER BY c.comment_date ASC",
            $start_date,
            $end_date
        );

        $notes = $wpdb->get_results($query);
        $vouchers = [];
		$order_ids = [];

        foreach ($notes as $note) {
			$order_ids[$note->order_id] = $note;
		}
		
        foreach ($order_ids as $order_id => $note) {
			$order = wc_get_order($order_id);
			if(get_class(wc_get_payment_gateway_by_order( $order )) !== 'WC_Gateway_COD'){
				continue;
			}
            if (preg_match('/\b\d{10}\b/', $note->note_content, $matches)) {
                $vouchers[$matches[0]] = [
                    'order_id' => $order_id,
                    'voucher_number' => $matches[0],
                    'voucher_date' => $note->voucher_date,
					'order_amount' => $order->get_total()
                ];
            }
        }

        return array_values($vouchers);
    }

    // Generate XLSX file
    private function generate_xlsx($data) {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Headers
        $sheet->setCellValue('A1', __('Order ID', 'geniki-voucher-export'));
        $sheet->setCellValue('B1', __('Voucher Number', 'geniki-voucher-export'));
        $sheet->setCellValue('C1', __('Voucher Date', 'geniki-voucher-export'));
        $sheet->setCellValue('D1', __('Order Amount', 'geniki-voucher-export'));

        // Data
        $row = 2;
        foreach ($data as $item) {
            $sheet->setCellValue('A' . $row, $item['order_id']);
            $sheet->setCellValue('B' . $row, $item['voucher_number']);
            $sheet->setCellValue('C' . $row, date('Y-m-d', strtotime($item['voucher_date'])));
            $sheet->setCellValue('D' . $row, $item['order_amount']);
            $row++;
        }

        // Format headers
        $sheet->getStyle('A1:C1')->getFont()->setBold(true);

        // Output
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="voucher_export.xlsx"');
        header('Cache-Control: max-age=0');

        $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;
    }
}

// Initialize plugin
new Geniki_Voucher_Export();

// Handle form submission
add_action('admin_init', function() {
    (new Geniki_Voucher_Export())->process_export();
});