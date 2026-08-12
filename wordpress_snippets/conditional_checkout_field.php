<?php
/**
 * Campo de checkout condicional por país (WooCommerce)
 * ----------------------------------------------------
 * Muestra el campo de código postal (obligatorio) solo para un país y lo oculta
 * para el resto. Pequeño ejemplo de personalización condicional del checkout.
 *
 * Snippet genérico y anonimizado.
 */

if ( ! defined( 'ABSPATH' ) ) { exit; }

add_filter( 'woocommerce_checkout_fields', 'conditional_postcode_field', 9999 );

function conditional_postcode_field( $fields ) {

    $target_country = 'ES'; // país que requiere el campo
    $country = '';

    if ( isset( $_POST['billing_country'] ) ) {
        $country = sanitize_text_field( wp_unslash( $_POST['billing_country'] ) );
    } elseif ( WC()->customer ) {
        $country = WC()->customer->get_billing_country();
    }

    if ( $country !== $target_country ) {
        unset( $fields['billing']['billing_postcode'] );
        unset( $fields['shipping']['shipping_postcode'] );
        return $fields;
    }

    foreach ( array( 'billing', 'shipping' ) as $section ) {
        $fields[ $section ][ $section . '_postcode' ] = array(
            'type'     => 'text',
            'label'    => __( 'Código postal', 'your-textdomain' ),
            'required' => true,
            'class'    => array( 'form-row-wide' ),
            'priority' => 110,
        );
    }

    return $fields;
}
