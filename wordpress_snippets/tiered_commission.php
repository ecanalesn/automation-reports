<?php
/**
 * Comisión de afiliado por tramos (AffiliateWP + WooCommerce)
 * ----------------------------------------------------------
 * Sobrescribe el importe de la comisión para que el porcentaje dependa del
 * total del pedido, por tramos:
 *
 *   total <= TIER_1  -> PCT_1
 *   total <= TIER_2  -> PCT_2
 *   total  > TIER_2  -> PCT_3
 *
 * Los umbrales se expresan en la divisa BASE del programa. Cuando un pedido se
 * cobra en otra divisa, un filtro permite inyectar una tasa de conversión.
 *
 * Snippet genérico y anonimizado: renombra las constantes/prefijo para tu proyecto.
 */

if ( ! defined( 'ABSPATH' ) ) { exit; }

if ( ! defined( 'AFF_TIER_1' ) ) define( 'AFF_TIER_1', 1000 );
if ( ! defined( 'AFF_TIER_2' ) ) define( 'AFF_TIER_2', 2000 );
if ( ! defined( 'AFF_PCT_1' ) )  define( 'AFF_PCT_1', 0.03 );
if ( ! defined( 'AFF_PCT_2' ) )  define( 'AFF_PCT_2', 0.025 );
if ( ! defined( 'AFF_PCT_3' ) )  define( 'AFF_PCT_3', 0.02 );

/**
 * Indica si este afiliado está sujeto a comisión por tramos.
 * Filtrable para poder excluir tipos concretos de afiliado.
 */
function aff_is_tiered_affiliate( $affiliate_id ) {
    return (bool) apply_filters( 'aff_is_tiered_affiliate', true, $affiliate_id );
}

/**
 * Total del pedido en la divisa base (importes de línea de producto, sin IVA ni envío).
 */
function aff_order_total_in_base( WC_Order $order ) {
    $total = 0.0;
    foreach ( $order->get_items() as $item ) {
        $total += (float) $item->get_total();
    }
    $base = get_option( 'woocommerce_currency' );
    $cur  = $order->get_currency();
    if ( $cur !== $base ) {
        $rate = apply_filters( 'aff_tier_exchange_rate', null, $cur, $base, $order );
        if ( $rate && (float) $rate > 0 ) {
            $total *= (float) $rate;
        }
    }
    return $total;
}

function aff_pct_for_total( $total ) {
    if ( $total <= AFF_TIER_1 ) return AFF_PCT_1;
    if ( $total <= AFF_TIER_2 ) return AFF_PCT_2;
    return AFF_PCT_3;
}

add_filter( 'affwp_calc_referral_amount', 'aff_tiered_referral_amount', 20, 5 );
function aff_tiered_referral_amount( $referral_amount, $affiliate_id = 0, $amount = 0, $reference = 0, $product_id = 0 ) {
    if ( ! function_exists( 'wc_get_order' ) ) return $referral_amount;
    $order = wc_get_order( $reference );
    if ( ! $order instanceof WC_Order ) return $referral_amount;
    if ( ! aff_is_tiered_affiliate( $affiliate_id ) ) return $referral_amount;

    $pct = aff_pct_for_total( aff_order_total_in_base( $order ) );
    $dec = function_exists( 'wc_get_price_decimals' ) ? wc_get_price_decimals() : 2;
    return round( (float) $amount * $pct, $dec );
}
