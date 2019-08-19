<?php
/**
 * Module2.bas
 * Option Explicit
 * TODO: remove all useless comments and DocBlocks when done translating whole file.
 */

/**
 * Const epsilon As Double = 0.0001
 * @var float
 */
const epsilon = 0.0001;

/**
 * Const column_offset As Long = 11
 * @var integer
 */
const column_offset = 11;

/*
 * data declarations
 *
 * */

/**
 * Private Type item_type_data
 * Class item_type_data
 */
class item_type_data
{
    /**
     * id As Long
     * @var integer
     */
    public $id;

    /**
     * width As Double
     * @var float
     */
    public $width;

    /**
     * height As Double
     * @var float
     */
    public $height;

    /**
     * length As Double
     * @var float
     */
    public $length;

    /**
     * volume As Double
     * @var float
     */
    public $volume;

    /**
     * weight As Double
     * @var float
     */
    public $weight;

    /**
     * xy_rotatable As Boolean
     * @var boolean
     */
    public $xy_rotatable;

    /**
     * yz_rotatable As Boolean
     * @var boolean
     */
    public $yz_rotatable;

    /**
     * heavy As Boolean
     * @var boolean
     */
    public $heavy;

    /**
     * fragile As Boolean
     * @var boolean
     */
    public $fragile;

    /**
     * mandatory As Long
     * @var integer
     */
    public $mandatory;

    /**
     * profit As Double
     * @var float
     */
    public $profit;

    /**
     * number_requested As Long
     * @var integer
     */
    public $number_requested;

    /**
     * sort_criterion As Double
     * @var float
     */
    public $sort_criterion;
}

/**
 * Private Type item_list_data
 * Class item_list_data
 */
class item_list_data
{
    /**
     * num_item_types As Long
     * @var integer
     */
    public $num_item_types;

    /**
     * total_number_of_items As Long
     * @var integer
     */
    public $total_number_of_items;

    /**
     * item_types() As item_type_data
     * @var item_type_data[]
     */
    public $item_types;

    // $item_types; is originally declared as item_types() As item_type_data
    // this means that it's a dynamic array of item_type_data objects.
    // this is where a relationship would help -> one to many
}

/**
 * Dim item_list As item_list_data
 * @var item_list_data
 */
$item_list = new item_list_data();

/**
 * Private Type container_type_data
 * Class container_type_data
 */
class container_type_data
{
    /**
     * type_id As Long
     * @var integer
     */
    public $type_id;

    /**
     * width As Double
     * @var float
     */
    public $width;

    /**
     * height As Double
     * @var float
     */
    public $height;

    /**
     * length As Double
     * @var float
     */
    public $length;

    /**
     * volume_capacity As Double
     * @var float
     */
    public $volume_capacity;

    /**
     * weight_capacity As Double
     * @var float
     */
    public $weight_capacity;

    /**
     * mandatory As Long
     * @var integer
     */
    public $mandatory;

    /**
     * cost As Double
     * @var float
     */
    public $cost;

    /**
     * number_available As Long
     * @var integer
     */
    public $number_available;
}

/**
 * Private Type container_list_data
 * Class container_list_data
 */
class container_list_data
{
    /**
     * num_container_types As Long
     * @var integer
     */
    public $num_container_types;

    /**
     * container_types() As container_type_data
     * @var container_type_data[]
     */
    public $container_types;

    // $container_types is originally declared as container_types() As container_type_data
    // this means that it's a dynamic array of container_type_data objects.
    // See if you can make collections from these things. (would be very helpful, because it
    // has many helper methods).
}

/**
 * Dim compatibility_list As compatibility_data
 * @var container_type_data
 */
$container_list = new container_list_data();

/**
 * Private Type compatibility_data
 * Class compatibility_data
 */
class compatibility_data
{
    /**
     * item_to_item() As Boolean
     * @var boolean[]
     */
    public $item_to_item;

    /**
     * container_to_item() As Boolean
     * @var boolean[]
     */
    public $container_to_item;

    // both $item_to_item and $container_to_item are declared originally as
    // 			item_to_item() As Boolean
    // 			container_to_item() As Boolean
    // this means that they are both dynamic arrays, containing boolean values.
    // they both look like a bidimensional arrays (read ahead in the original code).
}

/**
 * Dim compatibility_list As compatibility_data
 * @var compatibility_data
 */
$compatibility_list = new compatibility_data();

/**
 * Private Type item_location
 * Class item_location
 */
class item_location
{
    /**
     * origin_x As Double
     * @var float
     */
    public $origin_x;

    /**
     * origin_y As Double
     * @var float
     */
    public $origin_y;

    /**
     * origin_z As Double
     * @var float
     */
    public $origin_z;

    /**
     * next_to_item_type As Long
     * @var integer
     */
    public $next_to_item_type;
}

/**
 * Private Type item_in_container
 * Class item_in_container
 */
class item_in_container
{
    /**
     * item_type As Long
     * @var integer
     */
    public $item_type;

    /**
     * rotation As Long '1 to 6
     * @var integer
     */
    public $rotation;

    /**
     * mandatory As Long
     * @var integer
     */
    public $mandatory;

    /**
     * origin_x As Double
     * @var float
     */
    public $origin_x;

    /**
     * origin_y As Double
     * @var float
     */
    public $origin_y;

    /**
     * origin_z As Double
     * @var float
     */
    public $origin_z;

    /**
     * opposite_x As Double
     * @var float
     */
    public $opposite_x;

    /**
     * opposite_y As Double
     * @var float
     */
    public $opposite_y;

    /**
     * opposite_z As Double
     * @var float
     */
    public $opposite_z;
}

/**
 * Private Type container_data
 * Class container_data
 */
class container_data
{
    /**
     * type_id As Long
     * @var integer
     */
    public $type_id;

    /**
     * width As Double
     * @var float
     */
    public $width;

    /**
     * height As Double
     * @var float
     */
    public $height;

    /**
     * length As Double
     * @var float
     */
    public $length;

    /**
     * volume_capacity As Double
     * @var float
     */
    public $volume_capacity;

    /**
     * weight_capacity As Double
     * @var float
     */
    public $weight_capacity;

    /**
     * cost As Double
     * @var float
     */
    public $cost;

    /**
     * item_cnt As Long
     * @var integer
     */
    public $item_cnt;

    /**
     * mandatory As Long
     * @var integer
     */
    public $mandatory;

    /**
     * items() As item_in_container
     * @var item_in_container[]
     */
    public $items;

    /**
     * addition_point_count As Long
     * @var integer
     */
    public $addition_point_count;

    /**
     * addition_points() As item_location
     * @var item_location[]
     */
    public $addition_points;

    /**
     * repack_item_count() As Long
     * @var integer[]
     */
    public $repack_item_count;

    /**
     * volume_packed As Double
     * @var float
     */
    public $volume_packed;

    /**
     * weight_packed As Double
     * @var float
     */
    public $weight_packed;
}

/**
 * Private Type solution_data
 * Class solution_data
 */
class solution_data
{
    /**
     * num_containers As Long
     * @var integer
     */
    public $num_containers;

    /**
     * feasible As Boolean
     * @var boolean
     */
    public $feasible;

    /**
     * net_profit As Double
     * @var float
     */
    public $net_profit;

    /**
     * total_volume As Double
     * @var float
     */
    public $total_volume;

    /**
     * total_weight As Double
     * @var float
     */
    public $total_weight;

    /**
     * total_dispersion As Double
     * @var float
     */
    public $total_dispersion;

    /**
     * total_distance As Double
     * @var float
     */
    public $total_distance;

    /**
     * total_x_moment As Double
     * @var float
     */
    public $total_x_moment;

    /**
     * total_yz_moment As Double
     * @var float
     */
    public $total_yz_moment;

    /**
     * rotation_order() As Long
     * @var integer[]
     */
    public $rotation_order;

    /**
     * item_type_order() As Long
     * @var integer[]
     */
    public $item_type_order;

    /**
     * container() As container_data
     * @var container_data[]
     */
    public $container;

    /**
     * unpacked_item_count() As Long
     * @var integer[]
     */
    public $unpacked_item_count;
}

/**
 * Private Type instance_data
 * Class instance_data
 */
class instance_data
{
    /**
     * front_side_support As Boolean
     * @var boolean
     */
    public $front_side_support;

    /**
     * item_item_compatibility_worksheet As Boolean 'true if the data exists
     * @var boolean
     */
    public $item_item_compatibility_worksheet;

    /**
     * container_item_compatibility_worksheet As Boolean 'true if the data exists
     * @var boolean
     */
    public $container_item_compatibility_worksheet;
}

/**
 * Dim instance As instance_data
 * @var instance_data
 */
$instance = new instance_data();

/**
 * Private Type solver_option_data
 * Class solver_option_data
 */
class solver_option_data
{
    /**
     * CPU_time_limit As Double
     * @var float
     */
    public $CPU_time_limit;

    /**
     * item_sort_criterion As Long
     * @var integer
     */
    public $item_sort_criterion;

    /**
     * wall_building As Boolean
     * @var boolean
     */
    public $wall_building;
}

/**
 * Dim solver_options As solver_option_data
 * @var solver_option_data
 */
$solver_options = new solver_option_data();

/**
 * Private Type candidate_data
 * Class candidate_data
 */
class candidate_data
{
    /**
     * mandatory As Long
     * @var integer
     */
    public $mandatory;

    /**
     * net_profit As Double
     * @var float
     */
    public $net_profit;

    /**
     * total_volume As Double
     * @var float
     */
    public $total_volume;

    /**
     * item_type_to_be_added As Long
     * @var integer
     */
    public $item_type_to_be_added;
}

/**
 * helper for Rnd in VBA
 * @return float|int
 */
function random()
{
    return mt_rand() / mt_getrandmax();
}

/**
 * Private Sub SortContainers(solution As solution_data, random_reorder_probability As Double)
 * @param solution_data $solution
 * @param float $random_reorder_probability
 */
function SortContainers(solution_data $solution, float $random_reorder_probability)
{
    /*
     * TODO:
     * There were two With statements in the original code. Try to combine
     * them in a single one.
     * */
    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim candidate_index As Long
     * @var integer
     */
    $candidate_index = 0;

    /**
     * Dim max_mandatory As Long
     * @var integer
     */
    $max_mandatory = 0;

    /**
     * Dim max_volume_packed As Double
     * @var float
     */
    $max_volume_packed = 0.0;

    /**
     * Dim min_ratio As Double
     * @var float
     */
    $min_ratio = 0.0;

    /**
     * Dim swap_container As container_data
     * @var container_data
     */
    $swap_container = new container_data();

    /*
     * 'insertion sort
     *
     * */

    /**
     * If Rnd < 1 - random_reorder_probability Then
     */
    if (random() < 1 - $random_reorder_probability) {
        /*
         * 'insertion sort
         *
         * */

        /**
         * With solution
         * @var solution_data
         */
        $s = $solution;

        /**
         * For i = 1 To .num_contain
         * the . is actually the solution
         * since we're in the With statement
         */
        for ($i = 1; $i <= $s->num_containers; ++$i) {
            $candidate_index = $i;
            $max_mandatory = $s->container[$i]->mandatory;
            $max_volume_packed = $s->container[$i]->volume_packed;
            $min_ratio = $s->container[$i]->cost / $s->container[$i]->volume_capacity;

            /**
             * For j = i + 1 To .num_containers
             */
            for ($j = $i + 1; $j <= $s->num_containers; ++$j) {

                /**
                 * If (.container(j).mandatory > max_mandatory) Or _
                 *    ((.container(j).mandatory = max_mandatory) And (.container(j).volume_packed > max_volume_packed + epsilon)) Or _
                 *    ((.container(j).mandatory = 0) And (max_mandatory = 0) And (.container(j).volume_packed > max_volume_packed - epsilon) And ((.container(j).cost / .container(j).volume_capacity) < min_ratio)) Then
                 *    candidate_index = j
                 *    max_mandatory = .container(j).mandatory
                 *    max_volume_packed = .container(j).volume_packed
                 *    min_ratio = .container(j).cost / .container(j).volume_capacity
                 * End If
                 */
                if ($s->container[$j]->mandatory > $max_mandatory ||
                    ($s->container[$j]->mandatory == $max_mandatory && $s->container[$j]->volume_packed > $max_volume_packed + epsilon) ||
                    ($s->container[$j]->mandatory == 0 && $max_mandatory == 0 && $s->container[$j]->volume_packed > $max_volume_packed - epsilon && $s->container[$j]->cost / $s->container[$j]->volume_capacity < $min_ratio)) {
                    $candidate_index = $j;
                    $max_mandatory = $s->container[$j]->mandatory;
                    $max_volume_packed = $s->container[$j]->volume_packed;
                    $min_ratio = $s->container[$j]->cost / $s->container[$j]->volume_capacity;
                }
            }

            /**
             * If candidate_index <> i Then
             *     swap_container = .container(candidate_index)
             *     .container(candidate_index) = .container(i)
             *     .container(i) = swap_container
             * End If
             */
            if ($candidate_index != $i) {
                $swap_container = $s->container[$candidate_index];
                $s->container[$candidate_index] = $s->container[$i];
                $s->container[$i] = $swap_container;
            }
        }
    } else {
        /**
         * With solution
         * @var solution_data
         */
        $s = $solution;

        /**
         * For i = 1 To .num_containers
         *     candidate_index = Int((.num_containers - i + 1) * Rnd + i)
         *     If candidate_index <> i Then
         *         swap_container = .container(candidate_index)
         *         .container(candidate_index) = .container(i)
         *         .container(i) = swap_container
         *     End If
         * Next i
         */
        for ($i = 1; $i <= $s->num_containers; ++$i) {
            $candidate_index = intval(($s->num_containers - $i + 1) * random() + $i);
            if ($candidate_index != $i) {
                $swap_container = $s->container[$candidate_index];
                $s->container[$candidate_index] = $s->container[$i];
                $s->container[$i] = $swap_container;
            }
        }
    }
}

/**
 * Private Sub PerturbSolution(solution As solution_data, container_id As Long, percent_time_left As Double)
 *
 * * added item_list to param list to avoid dealing with globals
 *
 * @param solution_data $solution
 * @param int $container_id
 * @param float $percent_time_left
 * @param item_list_data $item_list
 */
function PerturbSolution(solution_data $solution, int $container_id, float $percent_time_left, item_list_data $item_list)
{
    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim k As Long
     * @var integer
     */
    $k = 0;

    /**
     * Dim l As Long
     * @var integer
     */
    $l = 0;

    /**
     * Dim swap_long As Long
     * @var integer
     */
    $swap_long = 0;

    /**
     * Dim operator_selection As Double
     * @var float
     */
    $operator_selection = 0.0;

    /**
     * Dim container_emptying_probability As Double
     * @var float
     */
    $container_emptying_probability = 0.0;

    /**
     * Dim item_removal_probability As Double
     * @var float
     */
    $item_removal_probability = 0.0;

    /**
     * Dim repack_flag As Boolean
     * @var boolean
     */
    $repack_flag = false;

    /**
     * Dim continue_flag As Boolean
     * @var boolean
     */
    $continue_flag = false;

    /**
     * Dim max_z As Double
     * @var float
     */
    $max_z = 0.0;

    /**
     * With solution.container(container_id)
     */
    $sc = $solution->container[$container_id];

    /*
     * '            container_emptying_probability = 1 - 0.8 * (.volume_packed / .volume_capacity)
     * '            item_removal_probability = 1 - 0.8 * (.volume_packed / .volume_capacity)
     *         'test
     * */

    /**
     * container_emptying_probability = 0.05 + 0.15 * percent_time_left
     * @var float
     */
    $container_emptying_probability = 0.05 + 0.15 * $percent_time_left;

    /**
     * item_removal_probability = 0.05 + 0.15 * percent_time_left
     * @var float
     */
    $item_removal_probability = 0.05 + 0.15 * $percent_time_left;

    /**
     * If .item_cnt > 0 Then
     */
    if ($sc->item_cnt > 0) {
        /**
         * If Rnd() < container_emptying_probability Then
         */
        if (random() < $container_emptying_probability) {
            /*
             * 'empty the container
             * */

            /**
             * For j = 1 To .item_cnt
             */
            for ($j = 1; $j <= $sc->item_cnt; ++$j) {
                /**
                 * solution.unpacked_item_count(.items(j).item_type) = solution.unpacked_item_count(.items(j).item_type) + 1
                 */
                $solution->unpacked_item_count[$sc->items[$j]->item_type] = $solution->unpacked_item_count[$sc->items[$j]->item_type] + 1;

                /*
                 * NOTE: had to add item_list to params if
                 * I didn't want to deal with global nonsense.
                 * */
                /**
                 * solution.net_profit = solution.net_profit - item_list.item_types(.items(j).item_type).profit
                 */
                $solution->net_profit = $solution->net_profit - $item_list->item_types[$sc->items[$j]->item_type]->profit;
            }

            /**
             * solution.net_profit = solution.net_profit + .cost
             */
            $solution->net_profit = $solution->net_profit + $sc->cost;

            /**
             * solution.total_volume = solution.total_volume - .volume_packed
             */
            $solution->total_volume = $solution->total_volume - $sc->volume_packed;

            /**
             * solution.total_weight = solution.total_weight - .weight_packed
             */
            $solution->total_weight = $solution->total_weight - $sc->weight_packed;

            /**
             * .item_cnt = 0
             */
            $sc->item_cnt = 0;

            /**
             * .volume_packed = 0
             */
            $sc->volume_packed = 0;

            /**
             * .weight_packed = 0
             */
            $sc->weight_packed = 0;

            /**
             * .addition_point_count = 1
             */
            $sc->addition_point_count = 1;

            /**
             * TODO: find out how this would translate.
             * .addition_points(1).origin_x = 0
             */
            $sc->addition_points[1]->origin_x = 0;

            /**
             * TODO: find out how this would translate.
             * .addition_points(1).origin_y = 0
             */
            $sc->addition_points[1]->origin_y = 0;

            /**
             * TODO: find out how this would translate.
             * .addition_points(1).origin_z = 0
             */
            $sc->addition_points[1]->origin_z = 0;

        } else {
            /**
             * repack_flag = False
             */
            $repack_flag = false;

            /**
             * TODO: noticed in the original code there's both Rnd and Rnd() -> find out what's the difference.
             * operator_selection = Rnd
             */
            $operator_selection = random();

            /**
             * If operator_selection < 0.3 Then
             */
            if ($operator_selection < 0.3) {
                /**
                 * For j = 1 To .item_cnt
                 */
                for ($j = 1; $j <= $sc->item_cnt; ++$j) {
                    /**
                     * If ((solution.feasible = False) And (.items(j).mandatory = 0)) Or (Rnd() < item_removal_probability) Then
                     */
                    if (($solution->feasible == false && $sc->items[$j]->mandatory == 0) || random() < $item_removal_probability) {
                        /**
                         * solution.unpacked_item_count(.items(j).item_type) = solution.unpacked_item_count(.items(j).item_type) + 1
                         */
                        $solution->unpacked_item_count[$sc->items[$j]->item_type] = $solution->unpacked_item_count[$sc->items[$j]->item_type] + 1;

                        /**
                         * solution.net_profit = solution.net_profit - item_list.item_types(.items(j).item_type).profit
                         */
                        $solution->net_profit = $solution->net_profit - $item_list->item_types[$sc->items[$j]->item_type]->profit;

                        /**
                         * .items(j).item_type = 0
                         */
                        $sc->items[$j]->item_type = 0;

                        /**
                         * repack_flag = True
                         */
                        $repack_flag = true;
                    }
                }
                /**
                 * ElseIf operator_selection < 0.3 Then
                 */
                /*
                 * TODO: Why are we checking again if operator_selection is < 0.3. Find a way to merge with original if
                 * */
            } else if ($operator_selection < 0.3) {
                /**
                 * max_z = 0
                 */
                $max_z = 0;

                /**
                 * For j = 1 To .item_cnt
                 */
                for ($j = 1; $j <= $sc->item_cnt; ++$j) {
                    /**
                     * If max_z < .items(j).opposite_z Then max_z = .items(j).opposite_z
                     */
                    if ($max_z < $sc->items[$j]->opposite_z) {
                        $max_z = $sc->items[$j]->opposite_z;
                    }
                }

                /**
                 * max_z = max_z * (0.1 + 0.5 * percent_time_left * Rnd)
                 */
                $max_z = $max_z * (0.1 + 0.5 * $percent_time_left * random());

                /**
                 * For j = 1 To .item_cnt
                 */
                for ($j = 1; $j <= $sc->item_cnt; ++$j) {
                    /**
                     * If ((solution.feasible = False) And (.items(j).mandatory = 0)) Or (.items(j).opposite_z < max_z) Then
                     */
                    if (($solution->feasible == false && $sc->items[$j]->mandatory == 0) || $sc->items[$j]->opposite_z < $max_z) {
                        /**
                         * solution.unpacked_item_count(.items(j).item_type) = solution.unpacked_item_count(.items(j).item_type) + 1
                         */
                        $solution->unpacked_item_count[$sc->items[$j]->item_type] = $solution->unpacked_item_count[$sc->items[$j]->item_type] + 1;

                        /**
                         * solution.net_profit = solution.net_profit - item_list.item_types(.items(j).item_type).profit
                         */
                        $solution->net_profit = $solution->net_profit - $item_list->item_types[$sc->items[$j]->item_type]->profit;

                        /**
                         * .items(j).item_type = 0
                         */
                        $sc->items[$j]->item_type = 0;

                        /**
                         * repack_flag = True
                         */
                        $repack_flag = true;
                    }
                }
            } else {
                /**
                 * max_z = 0
                 */
                $max_z = 0;

                /**
                 * For j = 1 To .item_cnt
                 */
                for ($j = 1; $j <= $sc->item_cnt; ++$j) {
                    /**
                     * If max_z < .items(j).opposite_z Then max_z = .items(j).opposite_z
                     */
                    if ($max_z < $sc->items[$j]->opposite_z) {
                        $max_z = $sc->items[$j]->opposite_z;
                    }
                }
                /**
                 * max_z = max_z * (0.6 - 0.5 * percent_time_left * Rnd)
                 */
                $max_z = $max_z * (0.6 - 0.5 * $percent_time_left * random());

                /**
                 * For j = 1 To .item_cnt
                 */
                for ($j = 1; $j <= $sc->item_cnt; ++$j) {
                    /**
                     * If ((solution.feasible = False) And (.items(j).mandatory = 0)) Or (.items(j).opposite_z > max_z) Then
                     */

                    if (($solution->feasible == false && $sc->items[$j]->mandatory == 0) || ($sc->items[$j]->opposite_z > $max_z)) {
                        /**
                         * solution.unpacked_item_count(.items(j).item_type) = solution.unpacked_item_count(.items(j).item_type) + 1
                         */
                        $solution->unpacked_item_count[$sc->items[$j]->item_type] = $solution->unpacked_item_count[$sc->items[$j]->item_type] + 1;

                        /**
                         * solution.net_profit = solution.net_profit - item_list.item_types(.items(j).item_type).profit
                         */
                        $solution->net_profit = $solution->net_profit - $item_list->item_types[$sc->items[$j]->item_type]->profit;

                        $sc->items[$j]->item_type = 0;
                        $repack_flag = true;
                    }
                }
            }
            /**
             * If repack_flag = True Then
             */
            if ($repack_flag == true) {
                /**
                 * For j = 1 To .item_cnt
                 */
                for ($j = 1; $j <= $sc->item_cnt; ++$j) {
                    /**
                     * If .items(j).item_type > 0 Then
                     */
                    if ($sc->items[$j]->item_type > 0) {
                        /**
                         * solution.net_profit = solution.net_profit - item_list.item_types(.items(j).item_type).profit
                         */
                        $solution->net_profit = $solution->net_profit - $item_list->item_types[$sc->items[$j]->item_type]->profit;
                    }
                }

                /**
                 * solution.net_profit = solution.net_profit + .cost
                 */
                $solution->net_profit = $solution->net_profit + $sc->cost;

                /**
                 * solution.total_volume = solution.total_volume - .volume_packed
                 */
                $solution->total_volume = $solution->total_volume - $sc->volume_packed;

                /**
                 * solution.total_weight = solution.total_weight - .weight_packed
                 */
                $solution->total_weight = $solution->total_weight - $sc->weight_packed;

                /**
                 * For j = 1 To item_list.num_item_types
                 */
                for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                    /**
                     * .repack_item_count(j) = 0
                     */
                    $sc->repack_item_count[$j] = 0;
                }

                /**
                 * For j = 1 To .item_cnt
                 */
                for ($j = 1; $j <= $sc->item_cnt; ++$j) {
                    /**
                     * If .items(j).item_type > 0 Then
                     */
                    if ($sc->items[$j]->item_type > 0) {
                        /**
                         * .repack_item_count(.items(j).item_type) = .repack_item_count(.items(j).item_type) + 1
                         */
                        $sc->repack_item_count[$sc->items[$j]->item_type] = $sc->repack_item_count[$sc->items[$j]->item_type] + 1;
                    }
                }

                /**
                 * .volume_packed = 0
                 */
                $sc->volume_packed = 0;

                /**
                 * .weight_packed = 0
                 */
                $sc->weight_packed = 0;

                /**
                 * .item_cnt = 0
                 */
                $sc->item_cnt = 0;

                /**
                 * .addition_point_count = 1
                 */
                $sc->addition_point_count = 1;

                /**
                 * .addition_points(1).origin_x = 0
                 */
                $sc->addition_points[1]->origin_x = 0;

                /**
                 * .addition_points(1).origin_y = 0
                 */
                $sc->addition_points[1]->origin_y = 0;

                /**
                 * .addition_points(1).origin_z = 0
                 */
                $sc->addition_points[1]->origin_z = 0;

                /*
                 * 'repack now
                 * */
                /**
                 * For j = 1 To item_list.num_item_types
                 */
                for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                    /**
                     * continue_flag = True
                     */
                    $continue_flag = true;
                    /**
                     * Do While (.repack_item_count(solution.item_type_order(j)) > 0) And (continue_flag = True)
                     */
                    do {
                        /**
                         * continue_flag = AddItemToContainer(solution, container_id, solution.item_type_order(j), 2, True)
                         */
                        $continue_flag = AddItemToContainer($solution, $container_id, $solution->item_type_order[$j], 2, true);
                    } while ($sc->repack_item_count[$solution->item_type_order[$j]] > 0 && $continue_flag == true);

                    /*
                     * ' put the remaining items in the unpacked items list
                     * */

                    /**
                     * solution.unpacked_item_count(solution.item_type_order(j)) = solution.unpacked_item_count(solution.item_type_order(j)) + .repack_item_count(solution.item_type_order(j))
                     */
                    $solution->unpacked_item_count[$solution->item_type_order[$j]] = $solution->unpacked_item_count[$solution->item_type_order[$j]] + $sc->repack_item_count[$solution->item_type_order[$j]];

                    /**
                     * .repack_item_count(solution.item_type_order(j)) = 0
                     */
                    $sc->repack_item_count[$solution->item_type_order[$j]] = 0;
                }
            }
        }
    }
}

/**
 * Private Sub PerturbRotationAndOrderOfItems(solution As solution_data)
 * @param solution_data $solution
 * @param item_list_data $item_list
 */
function PerturbRotationAndOrderOfItems(solution_data $solution, item_list_data $item_list)
{
    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim k As Long
     * @var integer
     */
    $k = 0;

    /**
     * Dim swap_long As Long
     * @var integer
     */
    $swap_long = 0;

    /*
     * 'change the preferred rotation order randomly
     * */

    /**
     * For i = 1 To item_list.num_item_types
     */
    for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
        /**
         * For j = 1 To 6
         */
        for ($j = 1; $j <= 6; ++$j) {
            /**
             * k = Int((6 - j + 1) * Rnd + j) ' the order to swap with
             */
            $k = intval((6 - $j + 1) * random() + $j);

            /**
             * swap_long = solution.rotation_order(i, j)
             */
            $swap_long = $solution->rotation_order[$i][$j];
            /*
             * TODO: find out if this is translated correctly.
             * rotation_order is integer[], so an array of integers.
             * but receiving two params, might mean it's a multidimensional array.
             * we'll just have to see.
             * */

            /**
             * solution.rotation_order(i, j) = solution.rotation_order(i, k)
             */
            $solution->rotation_order[$i][$j] = $solution->rotation_order[$i][$k];
            /*
             * TODO: recheck this after you know if the above one is translated correctly.
             * */

            /**
             * solution.rotation_order(i, k) = swap_long
             */
            $solution->rotation_order[$i][$k] = $swap_long;
            /*
             * TODO: check this as well.
             * */
        }
        /*
         * 'MsgBox ("Item type " & i & " rotation order: " & solution.rotation_order(i, 1) & solution.rotation_order(i, 2) & solution.rotation_order(i, 3) & solution.rotation_order(i, 4) & solution.rotation_order(i, 5) & solution.rotation_order(i, 6))
         * */
    }

    /*
     * 'change the item order randomly - test
     * */

    /**
     * For i = 1 To item_list.num_item_types
     */
    for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
        /**
         * j = Int((item_list.num_item_types - i + 1) * Rnd + i) ' the order to swap with
         */
        $j = intval(($item_list->num_item_types - $i + 1) * random() + $i);

        /**
         * swap_long = solution.item_type_order(i)
         */
        $swap_long = $solution->item_type_order[$i];

        /**
         * solution.item_type_order(i) = solution.item_type_order(j)
         */
        $solution->item_type_order[$i] = $solution->item_type_order[$j];

        /**
         * solution.item_type_order(j) = swap_long
         */
        $solution->item_type_order[$j] = $swap_long;
    }
}

/**
 * Private Function AddItemToContainer(solution As solution_data, container_index As Long, item_type_index As Long, add_type As Long, item_cohesion As Boolean)
 * * Added $instance to params to avoid globalization :))
 * * Added $compatibility_list to params
 * * Added $item_list to params
 * * Added $solver_options to params
 * @param solution_data $solution
 * @param int $container_index
 * @param int $item_type_index
 * @param int $add_type
 * @param bool $item_cohesion
 * @param instance_data $instance
 * @param compatibility_data $compatibility_list
 * @param item_list_data $item_list
 * @param solver_option_data $solver_options
 */
function AddItemToContainer(solution_data $solution, int $container_index, int $item_type_index, int $add_type, bool $item_cohesion, instance_data $instance, compatibility_data $compatibility_list, item_list_data $item_list, solver_option_data $solver_options)
{
    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim k As Long
     * @var integer
     */
    $k = 0;

    /**
     * Dim rotation_index As Long
     * @var integer
     */
    $rotation_index = 0;

    /**
     * Dim origin_x As Double
     * @var float
     */
    $origin_x = 0.0;

    /**
     * Dim origin_y As Double
     * @var float
     */
    $origin_y = 0.0;

    /**
     * Dim origin_z As Double
     * @var float
     */
    $origin_z = 0.0;

    /**
     * Dim opposite_x As Double
     * @var float
     */
    $opposite_x = 0.0;

    /**
     * Dim opposite_y As Double
     * @var float
     */
    $opposite_y = 0.0;

    /**
     * Dim opposite_z As Double
     * @var float
     */
    $opposite_z = 0.0;

    /**
     * Dim min_x As Double
     * @var float
     */
    $min_x = 0.0;

    /**
     * Dim min_y As Double
     * @var float
     */
    $min_y = 0.0;

    /**
     * Dim min_z As Double
     * @var float
     */
    $min_z = 0.0;

    /**
     * Dim next_to_item_type As Long
     * @var integer
     */
    $next_to_item_type = 0;

    /**
     * Dim candidate_position As Double
     * @var float
     */
    $candidate_position = 0.0;

    /**
     * Dim current_rotation As Long
     * @var integer
     */
    $current_rotation = 0;

    /**
     * Dim candidate_rotation As Double
     * @var float
     */
    $candidate_rotation = 0.0;

    /**
     * Dim area_supported as Double
     * @var float
     */
    $area_supported = 0.0;

    /**
     * Dim area_required as Double
     * @var float
     */
    $area_required = 0.0;

    /**
     * Dim intersection_right as Double
     * @var float
     */
    $intersection_right = 0.0;

    /**
     * Dim intersection_left as Double
     * @var float
     */
    $intersection_left = 0.0;

    /**
     * Dim intersection_top as Double
     * @var float
     */
    $intersection_top = 0.0;

    /**
     * Dim intersection_bottom as Double
     * @var float
     */
    $intersection_bottom = 0.0;

    /**
     * Dim support_flag as Boolean
     * @var boolean
     */
    $support_flag = false;

    /**
     * With solution.container(container_index)
     */
    $currentContainer = $solution->container[$container_index];

    /**
     * min_x = .width + 1
     */
    $min_x = $currentContainer->width + 1;

    /**
     * min_y = .height + 1
     */
    $min_y = $currentContainer->height + 1;

    /**
     * min_z = .length + 1
     */
    $min_z = $currentContainer->length + 1;

    /**
     * next_to_item_type = 0
     */
    $next_to_item_type = 0;

    /**
     * candidate_position = 0;
     */
    $candidate_position = 0;

    /*
     * 'compatibility check
     * */

    /**
     * If instance.container_item_compatibility_worksheet = True Then
     */
    if ($instance->container_item_compatibility_worksheet == true) {
        /**
         * If compatibility_list.container_to_item(.type_id, item_list.item_types(item_type_index).id) = False Then GoTo AddItemToContainer_Finish
         */
        if ($compatibility_list->container_to_item[$currentContainer->type_id][$item_list->item_types[$item_type_index]->id] == false) {
            /**
             * GoTo AddItemToContainer_Finish
             */
            $AddItemToContainer = AddItemToContainer_Finish($candidate_position, $solution, $container_index, $item_type_index, $candidate_rotation, $item_list, $add_type);
        }
    }

    /*
     * 'volume size check
     * */
    /**
     * If .volume_packed + item_list.item_types(item_type_index).volume > .volume_capacity Then GoTo AddItemToContainer_Finish
     */
    if ($currentContainer->volume_packed + $item_list->item_types[$item_type_index]->volume > $currentContainer->volume_capacity) {
        /**
         * GoTo AddItemToContainer_Finish
         */
        $AddItemToContainer = AddItemToContainer_Finish($candidate_position, $solution, $container_index, $item_type_index, $candidate_rotation, $item_list, $add_type);
    }

    /*
     * 'weight capacity check
     * */
    /**
     * If .weight_packed + item_list.item_types(item_type_index).weight > .weight_capacity Then GoTo AddItemToContainer_Finish
     */
    if ($currentContainer->weight_packed + $item_list->item_types[$item_type_index]->weight > $currentContainer->weight_capacity) {
        /**
         * GoTo AddItemToContainer_Finish
         */
        $AddItemToContainer = AddItemToContainer_Finish($candidate_position, $solution, $container_index, $item_type_index, $candidate_rotation, $item_list, $add_type);
    }

    /*
     * 'item to item compatibility check
     * */
    /**
     * If instance.item_item_compatibility_worksheet = True Then
     */
    if ($instance->item_item_compatibility_worksheet == true) {
        /**
         * For i = 1 To .item_cnt
         */
        for ($i = 1; $i <= $currentContainer->item_cnt; ++$i) {
            /**
             * If compatibility_list.item_to_item(item_list.item_types(item_type_index).id, item_list.item_types(.items(i).item_type).id) = False Then GoTo AddItemToContainer_Finish
             */
            if ($compatibility_list->item_to_item[$item_list->item_types[$item_type_index]->id][$item_list->item_types[$currentContainer->items[$i]->item_type]->id] == false) {
                /**
                 * GoTo AddItemToContainer_Finish
                 */
                $AddItemToContainer = AddItemToContainer_Finish($candidate_position, $solution, $container_index, $item_type_index, $candidate_rotation, $item_list, $add_type);
            }
        }
    }

    /**
     * For rotation_index = 1 To 6
     */

    for ($rotation_index = 1; $rotation_index <= 6; ++$rotation_index) {
        /*
         * 'test
         * 'If candidate_position <> 0 Then GoTo AddItemToContainer_Finish
         * */

        /**
         * current_rotation = solution.rotation_order(item_type_index, rotation_index)
         */
        $current_rotation = $solution->rotation_order[$item_type_index][$rotation_index];

        /*
         * 'forbidden rotations
         * */

        /**
         * If ((current_rotation = 3) Or (current_rotation = 4)) And (item_list.item_types(item_type_index).xy_rotatable = False) Then
         */
        if (($current_rotation == 3 || $current_rotation == 4) && $item_list->item_types[$item_type_index]->xy_rotatable == false) {
            /**
             * GoTo next_rotation_iteration
             */
            /*
             * TODO: find out a better way to write this, because next_rotation_iteration is just a continue statement.
             * */
            continue;
        }

        /**
         * If ((current_rotation = 5) Or (current_rotation = 6)) And (item_list.item_types(item_type_index).yz_rotatable = False) Then
         */
        if (($current_rotation == 5 || $current_rotation == 6) && $item_list->item_types[$item_type_index]->yz_rotatable == false) {
            /**
             * GoTo next_rotation_iteration
             */
            /*
             * TODO: find out a better way to write this, because next_rotation_iteration is just a continue statement.
             * */
            continue;
        }

        /*
         * 'symmetry breaking
         * */

        /**
         *  If (current_rotation = 2) And (Abs(item_list.item_types(item_type_index).width - item_list.item_types(item_type_index).length) < epsilon) Then
         */
        if ($current_rotation == 2 && abs($item_list->item_types[$item_type_index]->width - $item_list->item_types[$item_type_index]->length) < epsilon) {
            /**
             * GoTo next_rotation_iteration
             */
            /*
             * TODO: find out a better way to write this, because next_rotation_iteration is just a continue statement.
             * */
            continue;
        }

        /**
         * If (current_rotation = 4) And (Abs(item_list.item_types(item_type_index).width - item_list.item_types(item_type_index).height) < epsilon) Then
         */
        if ($current_rotation == 4 && abs($item_list->item_types[$item_type_index]->width - $item_list->item_types[$item_type_index]->height) < epsilon) {
            /**
             * GoTo next_rotation_iteration
             */
            /*
             * TODO: find out a better way to write this, because next_rotation_iteration is just a continue statement.
             * */
            continue;
        }

        /**
         * If (current_rotation = 6) And (Abs(item_list.item_types(item_type_index).height - item_list.item_types(item_type_index).length) < epsilon) Then
         */
        if (
            $current_rotation == 6 &&
            abs($item_list->item_types[$item_type_index]->height - $item_list->item_types[$item_type_index]->length) < epsilon
        ) {
            /**
             * GoTo next_rotation_iteration
             */
            /*
             * TODO: find out a better way to write this, because next_rotation_iteration is just a continue statement.
             * */
            continue;
        }

        /**
         * For i = 1 To .addition_point_count
         */
        for ($i = 1; $i <= $currentContainer->addition_point_count; ++$i) {

            /**
             * If (item_cohesion = True) And (candidate_position <> 0) And (next_to_item_type = item_type_index) And (.addition_points(i).next_to_item_type <> item_type_index) Then GoTo next_iteration
             */
            if ($item_cohesion == true && $candidate_position != 0 && $next_to_item_type == $item_type_index && $currentContainer->addition_points[$i]->next_to_item_type != $item_type_index) {
                /**
                 * GoTo next_iteration
                 */
                /*
                 * TODO: find out a better way to write this, because next_iteration is just a continue statement.
                 * */
                continue;
            }

            /**
             * origin_x = .addition_points(i).origin_x
             */
            $origin_x = $currentContainer->addition_points[$i]->origin_x;

            /**
             * origin_y = .addition_points(i).origin_y
             */
            $origin_y = $currentContainer->addition_points[$i]->origin_y;

            /**
             * origin_z = .addition_points(i).origin_z
             */
            $origin_z = $currentContainer->addition_points[$i]->origin_z;

            /**
             * If (item_list.item_types(item_type_index).heavy = True) And (origin_y > epsilon) Then GoTo next_iteration ' heavy item cannot be placed on any other item
             */
            if ($item_list->item_types[$item_type_index]->heavy == true && $origin_y > epsilon) {
                /**
                 * GoTo next_iteration
                 */
                /*
                 * TODO: find out a better way to write this, because next_iteration is just a continue statement.
                 * */
                continue;
            }

            /**
             * If current_rotation = 1 Then
             */
            if ($current_rotation == 1) {
                /**
                 * opposite_x = origin_x + item_list.item_types(item_type_index).width
                 */
                $opposite_x = $origin_x + $item_list->item_types[$item_type_index]->width;

                /**
                 * opposite_y = origin_y + item_list.item_types(item_type_index).height
                 */
                $opposite_y = $origin_y + $item_list->item_types[$item_type_index]->height;

                /**
                 * opposite_z = origin_z + item_list.item_types(item_type_index).length
                 */
                $opposite_z = $origin_z + $item_list->item_types[$item_type_index]->length;
                /**
                 * ElseIf current_rotation = 2 Then
                 */
            } else if ($current_rotation == 2) {
                /**
                 * opposite_x = origin_x + item_list.item_types(item_type_index).length
                 */
                $opposite_x = $origin_x + $item_list->item_types[$item_type_index]->length;

                /**
                 * opposite_y = origin_y + item_list.item_types(item_type_index).height
                 */
                $opposite_y = $origin_y + $item_list->item_types[$item_type_index]->height;

                /**
                 * opposite_z = origin_z + item_list.item_types(item_type_index).width
                 */
                $opposite_z = $origin_z + $item_list->item_types[$item_type_index]->width;
                /**
                 * ElseIf current_rotation = 3 Then
                 */
            } else if ($current_rotation == 3) {
                /**
                 * opposite_x = origin_x + item_list.item_types(item_type_index).width
                 */
                $opposite_x = $origin_x + $item_list->item_types[$item_type_index]->width;

                /**
                 * opposite_y = origin_y + item_list.item_types(item_type_index).length
                 */
                $opposite_y = $origin_y + $item_list->item_types[$item_type_index]->width;

                /**
                 * opposite_z = origin_z + item_list.item_types(item_type_index).height
                 */
                $opposite_z = $origin_z + $item_list->item_types[$item_type_index]->height;
                /**
                 * ElseIf current_rotation = 4 Then
                 */
            } else if ($current_rotation == 4) {
                /**
                 * opposite_x = origin_x + item_list.item_types(item_type_index).height
                 */
                $opposite_x = $origin_x + $item_list->item_types[$item_type_index]->height;

                /**
                 * opposite_y = origin_y + item_list.item_types(item_type_index).length
                 */
                $opposite_y = $origin_y + $item_list->item_types[$item_type_index]->length;

                /**
                 * opposite_z = origin_z + item_list.item_types(item_type_index).width
                 */
                $opposite_z = $origin_z + $item_list->item_types[$item_type_index]->width;

                /**
                 * ElseIf current_rotation = 5 Then
                 */
            } else if ($current_rotation == 5) {
                /**
                 * opposite_x = origin_x + item_list.item_types(item_type_index).height
                 */
                $opposite_x = $origin_x + $item_list->item_types[$item_type_index]->height;

                /**
                 * opposite_y = origin_y + item_list.item_types(item_type_index).width
                 */
                $opposite_y = $origin_y + $item_list->item_types[$item_type_index]->width;

                /**
                 * opposite_z = origin_z + item_list.item_types(item_type_index).length
                 */
                $opposite_z = $origin_z + $item_list->item_types[$item_type_index]->length;

                /**
                 * ElseIf current_rotation = 6 Then
                 */
            } else if ($current_rotation == 6) {
                /**
                 * opposite_x = origin_x + item_list.item_types(item_type_index).length
                 */
                $opposite_x = $origin_x + $item_list->item_types[$item_type_index]->length;

                /**
                 * opposite_y = origin_y + item_list.item_types(item_type_index).width
                 */
                $opposite_y = $origin_y + $item_list->item_types[$item_type_index]->width;

                /**
                 * opposite_z = origin_z + item_list.item_types(item_type_index).height
                 */
                $opposite_z = $origin_z + $item_list->item_types[$item_type_index]->height;
            }

            /*
             * 'check the feasibility of all four corners, w.r.t to the other items
             * */
            /**
             * If (opposite_x > .width + epsilon) Or (opposite_y > .height + epsilon) Or (opposite_z > .length + epsilon) Then GoTo next_iteration
             */
            if ($opposite_x > $currentContainer->width + epsilon || $opposite_y > $currentContainer->height + epsilon || $opposite_z > $currentContainer->length + epsilon) {
                /**
                 * GoTo next_iteration
                 */
                /*
                 * TODO: find out a better way to write this, because next_iteration is just a continue statement.
                 * */
                continue;
            }

            /**
             * For j = 1 To .item_cnt
             *      If (opposite_x < .items(j).origin_x + epsilon) Or _
             *          (.items(j).opposite_x < origin_x + epsilon) Or _
             *          (opposite_y < .items(j).origin_y + epsilon) Or _
             *          (.items(j).opposite_y < origin_y + epsilon) Or _
             *          (opposite_z < .items(j).origin_z + epsilon) Or _
             *          (.items(j).opposite_z < origin_z + epsilon) Then
             *          'no conflict
             *      Else
             *          'conflict
             *          GoTo next_iteration
             *      End If
             * Next j
             */
            for ($j = 1; $j <= $currentContainer->item_cnt; ++$j) {
                if (
                    $opposite_x < $currentContainer->items[$j]->origin_x + epsilon ||
                    $currentContainer->items[$j]->opposite_x < $origin_x + epsilon ||

                    $opposite_y < $currentContainer->items[$j]->origin_y + epsilon ||
                    $currentContainer->items[$j]->opposite_y < $origin_y + epsilon ||

                    $opposite_z < $currentContainer->items[$j]->origin_z + epsilon ||
                    $currentContainer->items[$j]->opposite_z < $origin_z + epsilon
                ) {
                    /*
                     * 'no conflict
                     * */
                    /*
                     * TODO: this statement is empty, the else part is just continue, so the whole loop looks useless.
                     * */
                } else {
                    continue;
                }
            }

            /*
             * 'vertical support
             * */

            /**
             * If origin_y < epsilon Then
             */
            if ($origin_y < epsilon) {
                /**
                 * support_flag = True
                 */
                $support_flag = true;
                /**
                 * Else
                 */
            } else {

                /**
                 * area_supported = 0
                 */
                $area_supported = 0;

                /**
                 * area_required = ((opposite_x - origin_x) * (opposite_z - origin_z))
                 */
                $area_required = ($opposite_x - $origin_x) * ($opposite_z - $origin_z);

                /**
                 * support_flag = False
                 */
                $support_flag = false;

                /**
                 * For j = .item_cnt To 1 Step -1
                 */
                for ($j = $currentContainer->item_cnt; $j >= 1; --$j) {

                    /**
                     * If (Abs(origin_y - .items(j).opposite_y) < epsilon) Then
                     */
                    if (abs($origin_y - $currentContainer->items[$j]->opposite_y) < epsilon) {
                        /*
                         * 'check for intersection
                         * */

                        /**
                         * intersection_right = opposite_x
                         */
                        $intersection_right = $opposite_x;

                        /**
                         * If intersection_right > .items(j).opposite_x Then intersection_right = .items(j).opposite_x
                         */
                        if ($intersection_right > $currentContainer->items[$j]->opposite_x) {
                            $intersection_right = $currentContainer->items[$j]->opposite_x;
                        }

                        /**
                         * intersection_left = origin_x
                         */
                        $intersection_left = $origin_x;

                        /**
                         * If intersection_left < .items(j).origin_x Then intersection_left = .items(j).origin_x
                         */
                        if ($intersection_left < $currentContainer->items[$j]->origin_x) {
                            $intersection_left = $currentContainer->items[$j]->origin_x;
                        }

                        /**
                         * intersection_top = opposite_z
                         */
                        $intersection_top = $opposite_z;

                        /**
                         * If intersection_top > .items(j).opposite_z Then intersection_top = .items(j).opposite_z
                         */
                        if ($intersection_top > $currentContainer->items[$j]->opposite_z) {
                            $intersection_top = $currentContainer->items[$j]->opposite_z;
                        }

                        /**
                         * intersection_bottom = origin_z
                         */
                        $intersection_bottom = $origin_z;

                        /**
                         * If intersection_bottom < .items(j).origin_z Then intersection_bottom = .items(j).origin_z
                         */
                        if ($intersection_bottom < $currentContainer->items[$j]->origin_z) {
                            $intersection_bottom = $currentContainer->items[$j]->origin_z;
                        }

                        /**
                         * If (intersection_right > intersection_left) And (intersection_top > intersection_bottom) Then
                         */
                        if ($intersection_right > $intersection_left && $intersection_top > $intersection_bottom) {
                            /*
                             * 'check for fragile items
                             * */
                            /**
                             * If item_list.item_types(.items(j).item_type).fragile = True Then
                             */
                            if ($item_list->item_types[$currentContainer->items[$j]->item_type]->fragile == true) {
                                /**
                                 * GoTo next_iteration
                                 */
                                /*
                                 * TODO: find out a better way to write this, because next_iteration is just a continue statement.
                                 * */
                                continue;
                            } else {
                                /**
                                 * area_supported = area_supported + (intersection_right - intersection_left) * (intersection_top - intersection_bottom)
                                 */
                                $area_supported = $area_supported + ($intersection_right - $intersection_left) * ($intersection_top - $intersection_bottom);

                                /**
                                 * If area_supported > area_required - epsilon Then
                                 */
                                if ($area_supported > $area_required - epsilon) {
                                    /**
                                     * support_flag = True
                                     */
                                    $support_flag = true;
                                    /**
                                     * Exit For
                                     */
                                    /*
                                     * TODO: find out what Exit For means. I'm just gonna guess it means break;
                                     * */
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            /**
             * If support_flag = False Then GoTo next_iteration
             */
            if ($support_flag == false) {
                /**
                 * GoTo next_iteration
                 */
                continue;
            }

            /*
             * 'front side support
             * */
            /**
             * If instance.front_side_support = True Then
             */
            if ($instance->front_side_support == true) {
                /**
                 * If origin_z < epsilon Then
                 */
                if ($origin_z < epsilon) {
                    /**
                     * support_flag = True
                     */
                    $support_flag = true;
                } else {
                    /**
                     * area_supported = 0
                     */
                    $area_supported = 0;

                    /**
                     * area_required = ((opposite_x - origin_x) * (opposite_y - origin_y))
                     */
                    $area_required = (($opposite_x - $origin_x) * ($opposite_y - $origin_y));

                    /**
                     * support_flag = False
                     */
                    $support_flag = false;

                    /**
                     * For j = .item_cnt To 1 Step -1
                     */
                    for ($j = $currentContainer->item_cnt; $j >= 1; --$j) {
                        /**
                         * If (Abs(origin_z - .items(j).opposite_z) < epsilon) Then
                         */
                        if (abs($origin_z - $currentContainer->items[$j]->opposite_z) < epsilon) {
                            /*
                             * 'check for intersection
                             * */

                            /**
                             * intersection_right = opposite_x
                             */
                            $intersection_right = $opposite_x;

                            /**
                             * If intersection_right > .items(j).opposite_x Then intersection_right = .items(j).opposite_x
                             */
                            if ($intersection_right > $currentContainer->items[$j]->opposite_x) {
                                $intersection_right = $currentContainer->items[$j]->opposite_x;
                            }

                            /**
                             * intersection_left = origin_x
                             */
                            $intersection_left = $origin_x;

                            /**
                             * If intersection_left < .items(j).origin_x Then intersection_left = .items(j).origin_x
                             */
                            if ($intersection_left < $currentContainer->items[$j]->origin_x) {
                                $intersection_left = $currentContainer->items[$j]->origin_x;
                            }

                            /**
                             * intersection_top = opposite_y
                             */
                            $intersection_top = $opposite_y;

                            /**
                             * If intersection_top > .items(j).opposite_y Then intersection_top = .items(j).opposite_y
                             */
                            if ($intersection_top > $currentContainer->items[$j]->opposite_y) {
                                $intersection_top = $currentContainer->items[$j]->opposite_y;
                            }

                            /**
                             * intersection_bottom = origin_y
                             */
                            $intersection_bottom = $origin_y;

                            /**
                             * If intersection_bottom < .items(j).origin_y Then intersection_bottom = .items(j).origin_y
                             */
                            if ($intersection_bottom < $currentContainer->items[$j]->origin_y) {
                                $intersection_bottom = $currentContainer->items[$j]->origin_y;
                            }

                            /**
                             * If (intersection_right > intersection_left) And (intersection_top > intersection_bottom) Then
                             */
                            if ($intersection_right > $intersection_left && $intersection_top > $intersection_bottom) {

                                /**
                                 * area_supported = area_supported + (intersection_right - intersection_left) * (intersection_top - intersection_bottom)
                                 */
                                $area_supported = $area_supported + ($intersection_right - $intersection_left) * ($intersection_top - $intersection_bottom);

                                /**
                                 * If area_supported > area_required - epsilon Then
                                 */
                                if ($area_supported > $area_required - epsilon) {
                                    /**
                                     * support_flag = True
                                     */
                                    $support_flag = true;
                                    /**
                                     * Exit For
                                     */
                                    /*
                                     * TODO: find out what Exit For means. I'm just gonna guess it means break;
                                     * */
                                    break;
                                }
                            }
                        }
                    }
                }
            }

            /**
             * If support_flag = False Then GoTo next_iteration
             */
            if ($support_flag == false) {
                /**
                 * GoTo next_iteration
                 */
                continue;
            }

            /*
             * 'no conflicts at this point
             * */

            /**
             * If (item_cohesion = True) And (next_to_item_type <> item_type_index) And (.addition_points(i).next_to_item_type = item_type_index) Then
             */
            if ($item_cohesion == true && $next_to_item_type != $item_type_index && $currentContainer->addition_points[$i]->next_to_item_type == $item_type_index) {
                /**
                 * min_x = origin_x
                 */
                $min_x = $origin_x;

                /**
                 * min_y = origin_y
                 */
                $min_y = $origin_y;

                /**
                 * min_z = origin_z
                 */
                $min_z = $origin_z;

                /**
                 * candidate_position = i
                 */
                $candidate_position = $i;

                /**
                 * candidate_rotation = current_rotation
                 */
                $candidate_rotation = $current_rotation;

                /**
                 * next_to_item_type = .addition_points(i).next_to_item_type
                 */
                $next_to_item_type = $currentContainer->addition_points[$i]->next_to_item_type;
                /**
                 * Else
                 */
            } else {
                /**
                 * If solver_options.wall_building = True Then
                 */
                if ($solver_options->wall_building == true) {
                    /**
                     * If (origin_z < min_z) Or _
                     * ((origin_z <= min_z + epsilon) And (origin_y < min_y)) Or _
                     * ((origin_z <= min_z + epsilon) And (origin_y <= min_y + epsilon) And (origin_x < min_x)) Then 'Or _
                     * ((origin_z <= min_z + epsilon) And (origin_y <= min_y + epsilon) And (origin_x <= min_x + epsilon) And ((opposite_x > .width + epsilon) Or (opposite_y > .height + epsilon))) Then
                     */
                    /*
                     * TODO: clarify the 'Or _ on the third line. or clarify the Then and the last line. either way, clarify something.
                     * */
                    if (
                        $origin_z < $min_z ||
                        ($origin_z <= $min_z + epsilon && $origin_y < $min_y) ||
                        ($origin_z <= $min_z + epsilon && $origin_y <= $min_y + epsilon && $origin_x && $origin_x < $min_x) ||
                        /*
                         * This is the breaking point where clarification is needed.
                         * */
                        ($origin_z <= $min_z + epsilon && $origin_y <= $min_y + epsilon && $origin_x <= $min_x + epsilon && ($opposite_x > $currentContainer->width + epsilon || $opposite_y > $currentContainer->height + epsilon))
                    ) {
                        /**
                         * min_x = origin_x
                         */
                        $min_x = $origin_x;

                        /**
                         * min_y = origin_y
                         */
                        $min_y = $origin_y;

                        /**
                         * min_z = origin_z
                         */
                        $min_z = $origin_z;

                        /**
                         * candidate_position = i
                         */
                        $candidate_position = $i;

                        /**
                         * candidate_rotation = current_rotation
                         */
                        $candidate_rotation = $current_rotation;

                        /**
                         * next_to_item_type = .addition_points(i).next_to_item_type
                         */
                        $next_to_item_type = $currentContainer->addition_points[$i]->next_to_item_type;
                    }
                    /**
                     * Else
                     */
                } else {
                    /**
                     * If (origin_y < min_y) Or _
                     *      ((origin_y <= min_y + epsilon) And (origin_z < min_z)) Or _
                     *      ((origin_y <= min_y + epsilon) And (origin_z <= min_z + epsilon) And (origin_x < min_x)) Then 'Or _
                     *      ((origin_y <= min_y + epsilon) And (origin_x <= min_x + epsilon) And (origin_z <= min_z + epsilon) And ((opposite_x > .width + epsilon) Or (opposite_y > .height + epsilon))) Then
                     */
                    /*
                     * TODO: clarify the 'Or _ on the third line.
                     * */
                    if (
                        $origin_y < $min_y ||
                        ($origin_y <= $min_y + epsilon && $origin_z < $min_z) ||
                        ($origin_y <= $min_y + epsilon && $origin_z <= $min_z + epsilon && $origin_x < $min_x) ||
                        /*
                         * Clarification needed here.
                         * */
                        (
                            $origin_y <= $min_y + epsilon && $origin_x <= $min_x + epsilon && $origin_z <= $min_z + epsilon &&
                            (
                                $opposite_x > $currentContainer->width + epsilon || $opposite_y > $currentContainer->height + epsilon
                            )
                        )
                    ) {
                        /**
                         * min_x = origin_x
                         */
                        $min_x = $origin_x;

                        /**
                         * min_y = origin_y
                         */
                        $min_y = $origin_y;

                        /**
                         * min_z = origin_z
                         */
                        $min_z = $origin_z;

                        /**
                         * candidate_position = i
                         */
                        $candidate_position = $i;

                        /**
                         * candidate_rotation = current_rotation
                         */
                        $candidate_rotation = $current_rotation;

                        /**
                         * next_to_item_type = .addition_points(i).next_to_item_type
                         */
                        $next_to_item_type = $currentContainer->addition_points[$i]->next_to_item_type;
                    }
                }
            }
        }
    }
    return $AddItemToContainer;
}

/**
 * next_iteration:
 *              Next i
 */

/*
 *
 * This looks just like a simple continue.
 *
 * */

/**
 * next_rotation_iteration:
 *          Next rotation_index
 *
 *      End With
 */

/*
 * This looks like a simple continue, but the End With kinda throws me off.
 * TODO: clarify which With the statement above ends.
 * */

/*
 * this just looks like the label where the GoTo sends. I'll refactor this to a simple function.
 * */


/**
 * AddItemToContainer_Finish:
 * @param float $candidate_position
 * @param solution_data $solution
 * @param int $container_index
 * @param int $item_type_index
 * @param float $candidate_rotation
 * @param item_list_data $item_list
 * @param int $add_type
 */
function AddItemToContainer_Finish(float $candidate_position, solution_data $solution, int $container_index, int $item_type_index, float $candidate_rotation, item_list_data $item_list, int $add_type)
{
    /**
     * If candidate_position = 0 Then
     */
    if ($candidate_position == 0) {
        /**
         * AddItemToContainer = False
         */
        $AddItemToContainer = false;
        return $AddItemToContainer;
        /*
         * TODO: clarify how this variable is defined.
         * It's in an if, so outside it, it doesn't make any sense.
         * We'll see if the original code uses it elsewhere.
         * */
        /**
         * Else
         */
    } else {
        /**
         * With solution.container(container_index)
         */
        $currentContainer = $solution->container[$container_index];

        /**
         * .item_cnt = .item_cnt + 1
         */
        $currentContainer->item_cnt = $currentContainer->item_cnt + 1;

        /**
         * .items(.item_cnt).item_type = item_type_index
         */
        $currentContainer->items[$currentContainer->item_cnt]->item_type = $item_type_index;

        /**
         * .items(.item_cnt).origin_x = .addition_points(candidate_position).origin_x
         */
        $currentContainer->items[$currentContainer->item_cnt]->origin_x = $currentContainer->addition_points[$candidate_position]->origin_x;
        /*
         * TODO: inconsistency found in declaration: $candidate_position is supposed to be a float. Illegal array key type: float
         * Clarify this.
         * */

        /**
         * .items(.item_cnt).origin_y = .addition_points(candidate_position).origin_y
         */
        $currentContainer->items[$currentContainer->item_cnt]->origin_y = $currentContainer->addition_points[$candidate_position]->origin_y;
        /*
         * TODO: inconsistency found in declaration: $candidate_position is supposed to be a float. Illegal array key type: float
         * Clarify this.
         * */

        /**
         * .items(.item_cnt).origin_z = .addition_points(candidate_position).origin_z
         */
        $currentContainer->items[$currentContainer->item_cnt]->origin_z = $currentContainer->addition_points[$candidate_position]->origin_z;
        /*
         * TODO: inconsistency found in declaration: $candidate_position is supposed to be a float. Illegal array key type: float
         * Clarify this.
         * */

        /**
         * .items(.item_cnt).rotation = candidate_rotation
         */
        $currentContainer->items[$currentContainer->item_cnt]->rotation = $candidate_rotation;

        /**
         * .items(.item_cnt).mandatory = item_list.item_types(item_type_index).mandatory
         */
        $currentContainer->items[$currentContainer->item_cnt]->mandatory = $item_list->item_types[$item_type_index]->mandatory;

        /**
         * If candidate_rotation = 1 Then
         */
        if ($candidate_rotation == 1) {

            /**
             * .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).width
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_x = $currentContainer->items[$currentContainer->item_cnt]->origin_x + $item_list->item_types[$item_type_index]->width;

            /**
             * .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).height
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_y = $currentContainer->items[$currentContainer->item_cnt]->origin_y + $item_list->item_types[$item_type_index]->height;

            /**
             * .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).length
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_z = $currentContainer->items[$currentContainer->item_cnt]->origin_z + $item_list->item_types[$item_type_index]->length;

            /**
             * ElseIf candidate_rotation = 2 Then
             */
        } else if ($candidate_rotation == 2) {
            /**
             * .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).length
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_x = $currentContainer->items[$currentContainer->item_cnt]->origin_x + $item_list->item_types[$item_type_index]->length;

            /**
             * .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).height
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_y = $currentContainer->items[$currentContainer->item_cnt]->origin_y + $item_list->item_types[$item_type_index]->height;

            /**
             * items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).width
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_z = $currentContainer->items[$currentContainer->item_cnt]->origin_z + $item_list->item_types[$item_type_index]->width;
            /**
             * ElseIf candidate_rotation = 3 Then
             */
        } else if ($candidate_rotation == 3) {
            /**
             * .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).width
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_x = $currentContainer->items[$currentContainer->item_cnt]->origin_x + $item_list->item_types[$item_type_index]->width;
            /**
             * .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).length
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_y = $currentContainer->items[$currentContainer->item_cnt]->origin_y + $item_list->item_types[$item_type_index]->length;
            /**
             * .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).height
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_z = $currentContainer->items[$currentContainer->item_cnt]->origin_z + $item_list->item_types[$item_type_index]->height;
            /**
             * ElseIf candidate_rotation = 4 Then
             */
        } else if ($candidate_rotation == 4) {
            /**
             * .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).height
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_x = $currentContainer->items[$currentContainer->item_cnt]->origin_x + $item_list->item_types[$item_type_index]->height;

            /**
             * .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).length
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_y = $currentContainer->items[$currentContainer->item_cnt]->origin_y + $item_list->item_types[$item_type_index]->length;

            /**
             * .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).width
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_z = $currentContainer->items[$currentContainer->item_cnt]->origin_z + $item_list->item_types[$item_type_index]->width;

            /**
             * ElseIf candidate_rotation = 5 Then
             */
        } else if ($candidate_rotation == 5) {
            /**
             * .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).height
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_x = $currentContainer->items[$currentContainer->item_cnt]->origin_x + $item_list->item_types[$item_type_index]->height;

            /**
             * .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).width
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_y = $currentContainer->items[$currentContainer->item_cnt]->origin_y + $item_list->item_types[$item_type_index]->width;

            /**
             * .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).length
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_z = $currentContainer->items[$currentContainer->item_cnt]->origin_z + $item_list->item_types[$item_type_index]->length;
            /**
             * ElseIf candidate_rotation = 6 Then
             */
        } else if ($candidate_rotation == 6) {
            /**
             * .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).length
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_x = $currentContainer->items[$currentContainer->item_cnt]->origin_x + $item_list->item_types[$item_type_index]->length;

            /**
             * .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).width
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_y = $currentContainer->items[$currentContainer->item_cnt]->origin_y + $item_list->item_types[$item_type_index]->width;

            /**
             * .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).height
             */
            $currentContainer->items[$currentContainer->item_cnt]->opposite_z = $currentContainer->items[$currentContainer->item_cnt]->origin_z + $item_list->item_types[$item_type_index]->height;
        }

        /**
         * .volume_packed = .volume_packed + item_list.item_types(item_type_index).volume
         */
        $currentContainer->volume_packed = $currentContainer->volume_packed + $item_list->item_types[$item_type_index]->volume;

        /**
         * .weight_packed = .weight_packed + item_list.item_types(item_type_index).weight
         */
        $currentContainer->weight_packed = $currentContainer->weight_packed + $item_list->item_types[$item_type_index]->weight;

        /**
         * If add_type = 2 Then
         */
        if ($add_type == 2) {
            /**
             * .repack_item_count(item_type_index) = .repack_item_count(item_type_index) - 1
             */
            $currentContainer->repack_item_count[$item_type_index] = $currentContainer->repack_item_count[$item_type_index] - 1;
        }

        /*
         * 'update the addition points
         * */

        /**
         * For i = candidate_position To .addition_point_count - 1
         */
        for ($i = $candidate_position; $i <= $currentContainer->addition_point_count - 1; ++$i) {
            /**
             * .addition_points(i) = .addition_points(i + 1)
             */
            $currentContainer->addition_points[$i] = $currentContainer->addition_points[$i + 1];
            /*
             * TODO: still need to figure out if $candidate_position is a float, since $i would also be floaty :)
             * */
        }

        /**
         * .addition_point_count = .addition_point_count - 1
         */
        $currentContainer->addition_point_count = $currentContainer->addition_point_count - 1;

        /**
         * If (.items(.item_cnt).opposite_x < .width - epsilon) And (.items(.item_cnt).origin_y < .height - epsilon) And (.items(.item_cnt).origin_z < .length - epsilon) Then
         */
        if (
            $currentContainer->items[$currentContainer->item_cnt]->opposite_x < $currentContainer->width - epsilon &&
            $currentContainer->items[$currentContainer->item_cnt]->origin_y < $currentContainer->height - epsilon &&
            $currentContainer->items[$currentContainer->item_cnt]->origin_z < $currentContainer->length - epsilon
        ) {
            /**
             * .addition_point_count = .addition_point_count + 1
             */
            $currentContainer->addition_point_count = $currentContainer->addition_point_count + 1;

            /**
             * .addition_points(.addition_point_count).origin_x = .items(.item_cnt).opposite_x
             */
            $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_x = $currentContainer->items[$currentContainer->item_cnt]->opposite_x;

            /**
             * .addition_points(.addition_point_count).origin_y = .items(.item_cnt).origin_y
             */
            $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_y = $currentContainer->items[$currentContainer->item_cnt]->origin_y;

            /**
             * .addition_points(.addition_point_count).origin_z = .items(.item_cnt).origin_z
             */
            $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_z = $currentContainer->items[$currentContainer->item_cnt]->origin_z;

            /**
             * .addition_points(.addition_point_count).next_to_item_type = item_type_index
             */
            $currentContainer->addition_points[$currentContainer->addition_point_count]->next_to_item_type = $item_type_index;
        }

        /**
         * If item_list.item_types(item_type_index).fragile = False Then ' no addition point on top of fragile items
         */
        if ($item_list->item_types[$item_type_index]->fragile == false) {
            /**
             * If (.items(.item_cnt).origin_x < .width - epsilon) And (.items(.item_cnt).opposite_y < .height - epsilon) And (.items(.item_cnt).origin_z < .length - epsilon) Then
             */
            if (
                $currentContainer->items[$currentContainer->item_cnt]->origin_x < $currentContainer->width - epsilon &&
                $currentContainer->items[$currentContainer->item_cnt]->opposite_y < $currentContainer->height - epsilon &&
                $currentContainer->items[$currentContainer->item_cnt]->origin_z < $currentContainer->length - epsilon
            ) {
                /**
                 *.addition_point_count = .addition_point_count + 1
                 */
                $currentContainer->addition_point_count = $currentContainer->addition_point_count + 1;

                /**
                 * .addition_points(.addition_point_count).origin_x = .items(.item_cnt).origin_x
                 */
                $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_x = $currentContainer->items[$currentContainer->item_cnt]->origin_x;

                /**
                 * .addition_points(.addition_point_count).origin_y = .items(.item_cnt).opposite_y
                 */
                $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_y = $currentContainer->items[$currentContainer->item_cnt]->opposite_y;

                /**
                 * .addition_points(.addition_point_count).origin_z = .items(.item_cnt).origin_z
                 */
                $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_z = $currentContainer->items[$currentContainer->item_cnt]->origin_z;

                /**
                 * .addition_points(.addition_point_count).next_to_item_type = item_type_index
                 */
                $currentContainer->addition_points[$currentContainer->addition_point_count]->next_to_item_type = $item_type_index;
            }
        }

        /**
         * If (.items(.item_cnt).origin_x < .width - epsilon) And (.items(.item_cnt).origin_y < .height - epsilon) And (.items(.item_cnt).opposite_z < .length - epsilon) Then
         */
        if (
            $currentContainer->items[$currentContainer->item_cnt]->origin_x < $currentContainer->width - epsilon &&
            $currentContainer->items[$currentContainer->item_cnt]->origin_y < $currentContainer->height - epsilon &&
            $currentContainer->items[$currentContainer->item_cnt]->opposite_z < $currentContainer->length - epsilon
        ) {
            /**
             * .addition_point_count = .addition_point_count + 1
             */
            $currentContainer->addition_point_count = $currentContainer->addition_point_count + 1;

            /**
             * .addition_points(.addition_point_count).origin_x = .items(.item_cnt).origin_x
             */
            $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_x = $currentContainer->items[$currentContainer->item_cnt]->origin_x;

            /**
             * .addition_points(.addition_point_count).origin_y = .items(.item_cnt).origin_y
             */
            $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_y = $currentContainer->items[$currentContainer->item_cnt]->origin_y;

            /**
             * .addition_points(.addition_point_count).origin_z = .items(.item_cnt).opposite_z
             */
            $currentContainer->addition_points[$currentContainer->addition_point_count]->origin_z = $currentContainer->items[$currentContainer->item_cnt]->opposite_z;

            /**
             * .addition_points(.addition_point_count).next_to_item_type = item_type_index
             */
            $currentContainer->addition_points[$currentContainer->addition_point_count]->next_to_item_type = $item_type_index;
        }
        /**
         * End With
         */

        /**
         * With solution
         */
        $s = $solution;
        /*
         * 'update the profit
         * */

        /**
         * If .container(container_index).item_cnt = 1 Then
         */
        if ($s->container[$container_index]->item_cnt == 1) {
            /**
             * .net_profit = .net_profit + item_list.item_types(item_type_index).profit - .container(container_index).cost
             */
            $s->net_profit = $s->net_profit + $item_list->item_types[$item_type_index]->profit - $s->container[$container_index]->cost;
            /**
             * Else
             */
        } else {
            $s->net_profit = $s->net_profit + $item_list->item_types[$item_type_index]->profit;
        }

        /*
         * 'update the volume per container and the total volume
         * */

        /**
         * .total_volume = .total_volume + item_list.item_types(item_type_index).volume
         */
        $s->total_volume = $s->total_volume + $item_list->item_types[$item_type_index]->volume;

        /**
         * .total_weight = .total_weight + item_list.item_types(item_type_index).weight
         */
        $s->total_weight = $s->total_weight + $item_list->item_types[$item_type_index]->weight;

        /*
         * 'update the unpacked items
         * */

        /**
         * If add_type = 1 Then
         */
        if ($add_type == 1) {
            /**
             * .unpacked_item_count(item_type_index) = .unpacked_item_count(item_type_index) - 1
             */
            $s->unpacked_item_count[$item_type_index] = $s->unpacked_item_count[$item_type_index] - 1;
        }

        /**
         * End With
         */

        /**
         * AddItemToContainer = True
         */
        $AddItemToContainer = true;
        /*
         * TODO: clarify what this variable actually does.
         * We've set it at the top of the function as false.
         * Now it's true. Will it ever be used?
         * */
        return $AddItemToContainer;
    }
}


/**
 * Private Sub GetSolverOptions()
 * I'm guessing that this handles the data retrieval
 * from the worksheet.
 * This can be translated with additional params, which can later be retrieved from a frontend form.
 * @param solver_option_data $solver_options
 * @param string $item_sort_criterion
 * @param string $wall_building
 * @param int $CPU_time_limit
 */
function GetSolverOptions(solver_option_data $solver_options, string $item_sort_criterion, string $wall_building, int $CPU_time_limit)
{
    /**
     * ThisWorkbook.Worksheets("CLP Solver Console").Activate
     */
    /*
     * This makes sure that the worksheet is in focus, I'm guessing.
     * */

    /**
     * With solver_options
     * I already have $solver_options as a param.
     */

    /**
     * If Cells(12, 3).Value = "Volume" Then
     *      .item_sort_criterion = 1
     * ElseIf Cells(12, 3).Value = "Weight" Then
     *      .item_sort_criterion = 2
     * Else
     *      .item_sort_criterion = 3
     * End If
     *
     * In our case, we're getting this data through the $item_sort_criterion param.
     */
    if ($item_sort_criterion == "Volume") {
        $solver_options->item_sort_criterion = 1;
    } else if ($item_sort_criterion == "Weight") {
        $solver_options->item_sort_criterion = 2;
    } else {
        $solver_options->item_sort_criterion = 3;
    }

    /**
     * If Cells(13, 3).Value = "Wall-building" Then
     *      .wall_building = True
     * Else
     *      .wall_building = False
     * End If
     *
     * In our case, we're getting this data through the $wall_building param.
     */
    if ($wall_building == "Wall-building") {
        $solver_options->wall_building = true;
    } else {
        $solver_options->wall_building = false;
    }

    /**
     * .CPU_time_limit = Cells(14, 3).Value
     *
     * In our case, we're getting this data through the $CPU_time_limit param.
     */
    $solver_options->CPU_time_limit = $CPU_time_limit;

    /**
     * End With
     */
}

/**
 * Private Sub GetItemData()
 * This stores all the item data into the $item_list.
 * Aside from passing in the number of item types, we'll need
 * to also pass in additional information for all the items.
 *
 * A proper array of items would contain items of type item_type_data.
 *
 * Anyway, this whole func seems redundant, since we're just copying
 * everything from $items to $item_list.
 *
 *
 * @param item_list_data $item_list
 * @param int $num_item_types
 * @param item_type_data[] $items
 * @param solver_option_data $solver_options
 */
function GetItemData(item_list_data $item_list, int $num_item_types, array $items, solver_option_data $solver_options)
{
    /**
     * item_list.num_item_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(2, 3).Value
     *
     * We're going to transform the ThisWorkbook... thing into a param. We'll call it $num_item_types, for convenience.
     */
    $item_list->num_item_types = $num_item_types;

    /**
     * item_list.total_number_of_items = 0
     */
    $item_list->total_number_of_items = 0;

    /**
     * ReDim item_list.item_types(1 To item_list.num_item_types)
     *
     * I'm guessing ReDim is just redeclaring the variable to a new val.
     * In this case, it creates a range from 1 to $item_list->num_item_types
     */
    $item_list->item_types = range(1, $item_list->num_item_types);

    /**
     * ThisWorkbook.Worksheets("1.Items").Activate
     * This should bring the "1.Items" worksheet into focus.
     */

    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim max_volume As Double
     * @var float
     */
    $max_volume = 0.0;

    /**
     * Dim max_weight As Double
     * @var float
     */
    $max_weight = 0.0;

    /**
     * max_volume = 0
     * Already declared an initial value, no need to do it again.
     */

    /**
     * max_weight = 0
     * Already declared an initial value, no need to do it again.
     */

    /**
     * With item_list
     * Let's just call item_list as iL, in case we get really long lines again.
     */
    $iL = $item_list;

    /**
     * For i = 1 To .num_item_types
     */
    for ($i = 1; $i <= $iL->num_item_types; ++$i) {
        /**
         * .item_types(i).id = i
         */
        $iL->item_types[$i]->id = $i;

        /**
         * .item_types(i).width = Cells(2 + i, 4).Value
         * well, this gets specific widths from the table on the worksheet.
         * we'll need to pass in our dimensions for all the items.
         */
        $iL->item_types[$i]->width = $items[$i]->width;

        /**
         * .item_types(i).height = Cells(2 + i, 5).Value
         */
        $iL->item_types[$i]->height = $items[$i]->height;

        /**
         * .item_types(i).length = Cells(2 + i, 6).Value
         */
        $iL->item_types[$i]->length = $items[$i]->length;

        /**
         * .item_types(i).volume = Cells(2 + i, 7).Value
         *
         * This is calculated within the worksheet, so we'll move the calculation here.
         */
        $iL->item_types[$i]->volume = $items[$i]->width * $items[$i]->height * $items[$i]->length;

        /**
         * If max_volume < .item_types(i).volume Then
         */
        if ($max_volume < $iL->item_types[$i]->volume) {
            /**
             * max_volume = .item_types(i).volume
             */
            $max_volume = $iL->item_types[$i]->volume;
        }

        /**
         * If Cells(2 + i, 9).Value = "Yes" Then
         *      .item_types(i).xy_rotatable = True
         * Else
         *      .item_types(i).xy_rotatable = False
         * End If
         *
         * This is gotten from the cells, but we already have this data in the $items param.
         */
        if ($items[$i]->xy_rotatable == true) {
            $iL->item_types[$i]->xy_rotatable = true;
        } else {
            $iL->item_types[$i]->xy_rotatable = false;
        }

        /**
         * If Cells(2 + i, 10).Value = "Yes" Then
         *      .item_types(i).yz_rotatable = True
         * Else
         *      .item_types(i).yz_rotatable = False
         * End If
         *
         * This handles getting the request for yz_rotatable.
         * It gets overwritten if a certain condition (cubic?) happens.
         * There's no point in rotating a cube for a better fit :)
         */
        if ($items[$i]->yz_rotatable == true) {
            $iL->item_types[$i]->yz_rotatable = true;
        } else {
            $iL->item_types[$i]->yz_rotatable = false;
        }

        /**
         * If (Abs(.item_types(i).width - .item_types(i).height) < epsilon) And (Abs(.item_types(i).width - .item_types(i).length) < epsilon) Then
         *
         */
        if (abs($iL->item_types[$i]->width - $iL->item_types[$i]->height < epsilon) && abs($iL->item_types[$i]->width - $iL->item_types[$i]->length) < epsilon) {
            /**
             * .item_types(i).xy_rotatable = False
             */
            $iL->item_types[$i]->xy_rotatable = false;

            /**
             * .item_types(i).yz_rotatable = False
             */
            $iL->item_types[$i]->yz_rotatable = false;
        }

        /**
         * .item_types(i).weight = Cells(2 + i, 11).Value
         */
        $iL->item_types[$i]->weight = $items[$i]->weight;

        /**
         * If max_weight < .item_types(i).weight Then
         */
        if ($max_weight < $iL->item_types[$i]->weight) {
            /**
             * max_weight = .item_types(i).weight
             */
            $max_weight = $iL->item_types[$i]->weight;
        }

        /**
         * If Cells(2 + i, 12).Value = "Yes" Then
         *      .item_types(i).heavy = True
         * Else
         *      .item_types(i).heavy = False
         * End If
         *
         * We already carry this data in the $items array.
         * Of course, this entire if can be rewritten as:
         * * $il->item_types[$i]->heavy = $items[$i]->heavy
         * since both data types are the same -> boolean
         */
        if ($items[$i]->heavy == true) {
            $iL->item_types[$i]->heavy = true;
        } else {
            $iL->item_types[$i]->heavy = true;
        }

        /**
         * If Cells(2 + i, 13).Value = "Yes" Then
         *      .item_types(i).fragile = True
         * Else
         *      .item_types(i).fragile = False
         * End If
         *
         * We already have this data in the $items array.
         * Of course, this entire if can be rewritten as:
         * * $il->item_types[$i]->fragile = $items[$i]->fragile
         * since both data types are the same -> boolean
         */
        if ($items[$i]->fragile == true) {
            $iL->item_types[$i]->fragile = true;
        } else {
            $iL->item_types[$i]->fragile = false;
        }

        /**
         * If Cells(2 + i, 14).Value = "Must be packed" Then
         *      .item_types(i).mandatory = 1
         * ElseIf Cells(2 + i, 14).Value = "May be packed" Then
         *      .item_types(i).mandatory = 0
         * ElseIf Cells(2 + i, 14).Value = "Don't pack" Then
         *      .item_types(i).mandatory = -1
         * End If
         *
         * This will be a very redundant statement.
         */
        if ($items[$i]->mandatory == 1) {
            $iL->item_types[$i]->mandatory = 1;
        } else if ($items[$i]->mandatory == 0) {
            $iL->item_types[$i]->mandatory = 0;
        } else if ($items[$i]->mandatory == -1) {
            $iL->item_types[$i]->mandatory = -1;
        }
        /*
         * TODO: remember to encode ["Must be packed", "May be packed", "Don't pack"] to [1, 0, -1]
         * */

        /**
         * .item_types(i).profit = Cells(2 + i, 15).Value
         */
        $iL->item_types[$i]->profit = $items[$i]->profit;

        /**
         * .item_types(i).number_requested = Cells(2 + i, 16).Value
         */
        $iL->item_types[$i]->number_requested = $items[$i]->number_requested;

        /**
         * item_list.total_number_of_items = item_list.total_number_of_items + .item_types(i).number_requested
         */
        $item_list->total_number_of_items = $item_list->total_number_of_items + $iL->item_types[$i]->number_requested;
    }

    /**
     * For i = 1 To .num_item_types
     */
    for ($i = 1; $i <= $iL->num_item_types; ++$i) {
        /**
         * If solver_options.item_sort_criterion = 1 Then
         */
        if ($solver_options->item_sort_criterion == 1) {
            /**
             * .item_types(i).sort_criterion = .item_types(i).volume * (max_weight + 1) + .item_types(i).weight
             */
            $iL->item_types[$i]->sort_criterion = $iL->item_types[$i]->volume * ($max_weight + 1) + $iL->item_types[$i]->weight;
            /**
             * ElseIf solver_options.item_sort_criterion = 2 Then
             */
        } else if ($solver_options->item_sort_criterion == 2) {
            /**
             * .item_types(i).sort_criterion = .item_types(i).weight * (max_volume + 1) + .item_types(i).volume
             */
            $iL->item_types[$i]->sort_criterion = $iL->item_types[$i]->weight * ($max_volume + 1) + $iL->item_types[$i]->volume;
            /**
             * Else
             */
        } else {
            /**
             * .item_types(i).sort_criterion = .item_types(i).width
             */
            $iL->item_types[$i]->sort_criterion = $iL->item_types[$i]->width;

            /**
             * If .item_types(i).sort_criterion < .item_types(i).height Then
             */
            if ($iL->item_types[$i]->sort_criterion < $iL->item_types[$i]->height) {
                /**
                 * .item_types(i).sort_criterion = .item_types(i).height
                 */
                $iL->item_types[$i]->sort_criterion = $iL->item_types[$i]->weight;
            }

            /**
             * If .item_types(i).sort_criterion < .item_types(i).length Then
             */
            if ($iL->item_types[$i]->sort_criterion < $iL->item_types[$i]->length) {
                /**
                 * .item_types(i).sort_criterion = .item_types(i).length
                 */
                $iL->item_types[$i]->sort_criterion = $iL->item_types[$i]->length;
            }

            /**
             * .item_types(i).sort_criterion = .item_types(i).sort_criterion * (max_volume + 1) + .item_types(i).volume
             */
            $iL->item_types[$i]->sort_criterion = $iL->item_types[$i]->sort_criterion * ($max_volume + 1) + $iL->item_types[$i]->volume;
        }
    }
}

/**
 * Private Sub GetContainerData()
 *
 * Same as above function, just gets the data from
 * the worksheet and stores it. Redundant in our case.
 *
 * @param container_list_data $container_list
 * @param int $num_container_types
 * @param container_type_data[] $containers
 */

function GetContainerData(container_list_data $container_list, int $num_container_types, array $containers)
{
    /**
     * container_list.num_container_types = ThisWorkbook.Worksheets("CLP Solver Console").Cells(4, 3).Value
     *
     * this gets the actual number of containers from the setup sheet.
     * we'll just pass it in as a param in our case.
     */
    $container_list->num_container_types = $num_container_types;

    /**
     * ReDim container_list.container_types(1 To container_list.num_container_types)
     *
     * I'm guessing ReDim is just redeclaring the variable to a new val.
     * In this case, it creates a range from 1 to $container_list->num_container_types
     */
    $container_list->container_types = range(1, $container_list->num_container_types);

    /**
     * ThisWorkbook.Worksheets("2.Containers").Activate
     *
     * This just focuses the "2.Containers" worksheet, i guess.
     */

    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * With container_list
     *
     * We'll just short the $container_list to $cL
     */
    $cL = $container_list;

    /**
     * For i = 1 To .num_container_types
     */
    for ($i = 1; $i <= $cL->num_container_types; ++$i) {
        /**
         * .container_types(i).type_id = i
         */
        $cL->container_types[$i]->type_id = $i;

        /**
         * .container_types(i).width = Cells(1 + i, 3).Value
         *
         * Since we're passing in an array of $containers,
         * we don't have to bother with worksheets or other
         * exotics.
         */
        $cL->container_types[$i]->width = $containers[$i]->width;

        /**
         * .container_types(i).height = Cells(1 + i, 4).Value
         *
         * Same as above
         */
        $cL->container_types[$i]->height = $containers[$i]->height;

        /**
         * .container_types(i).length = Cells(1 + i, 5).Value
         *
         * Same as above
         */
        $cL->container_types[$i]->length = $containers[$i]->length;

        /**
         * .container_types(i).volume_capacity = Cells(1 + i, 6).Value
         */
        $cL->container_types[$i]->volume_capacity = $containers[$i]->volume_capacity;

        /**
         * .container_types(i).weight_capacity = Cells(1 + i, 7).Value
         */
        $cL->container_types[$i]->weight_capacity = $containers[$i]->weight_capacity;

        /**
         * If Cells(1 + i, 8).Value = "Must be used" Then
         *      .container_types(i).mandatory = 1
         * ElseIf Cells(1 + i, 8).Value = "May be used" Then
         *      .container_types(i).mandatory = 0
         * ElseIf Cells(1 + i, 8).Value = "Do not use" Then
         *      .container_types(i).mandatory = -1
         * End If
         */
        /*
         * TODO: remember to encode ["Must be used", "May be used", "Do not use"] to [1, 0, -1]
         * */
        if ($containers[$i]->mandatory == 1) {
            $cL->container_types[$i]->mandatory = 1;
        } else if ($containers[$i]->mandatory == 0) {
            $cL->container_types[$i]->mandatory = 0;
        } else if ($containers[$i]->mandatory == -1) {
            $cL->container_types[$i]->mandatory = -1;
        }

        /**
         * .container_types(i).cost = Cells(1 + i, 9).Value
         */
        $cL->container_types[$i]->cost = $containers[$i]->cost;

        /**
         * .container_types(i).number_available = Cells(1 + i, 10).Value
         */
        $cL->container_types[$i]->number_available = $containers[$i]->number_available;
    }
}

/**
 * Private Sub GetCompatibilityData()
 *
 * @param compatibility_data $compatibility_list
 * @param instance_data $instance
 * @param item_list_data $item_list
 * @param container_list_data $container_list
 * @param array $item_item_compatibilities
 * @param array $container_item_compatibilities
 */
function GetCompatibilityData(compatibility_data $compatibility_list, instance_data $instance, item_list_data $item_list, container_list_data $container_list, array $item_item_compatibilities, array $container_item_compatibilities)
{
    /**
     * With compatibility_list
     *
     * We'll just shorten $compatibility_list to $cL
     */
    $cL = $compatibility_list;

    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim k As Long
     * @var integer
     */
    $k = 0;

    /**
     * If instance.item_item_compatibility_worksheet = True Then
     */
    if ($instance->item_item_compatibility_worksheet == true) {
        /**
         * ReDim .item_to_item(1 To item_list.num_item_types, 1 To item_list.num_item_types)
         *
         * This means that we're redefining the $cL->item_to_item to two ranges -> two dimensional array.
         * We'll start simple, with an array definition.
         * I'm guessing that ReDim just specifies the dimensions in advance. We'll see how this gets filled.
         */
        $cL->item_to_item = array();

        /**
         * For i = 1 To item_list.num_item_types
         */
        for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
            /**
             * For j = 1 To item_list.num_item_types
             */
            for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                /**
                 * .item_to_item(i, j) = True
                 */
                $cL->item_to_item[$i][$j] = true;
                /**
                 * I guess now we know the purpose of the ReDim above.
                 * This just fills the two dimensional array with
                 * data.
                 */
            }
        }

        /**
         * k = 3
         *
         * I wonder what the purpose of the hardcoded 3 is...
         */
        $k = 3;

        /**
         * For i = 1 To item_list.num_item_types
         */
        for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
            /**
             * For j = i + 1 To item_list.num_item_types
             */
            for ($j = $i + 1; $j <= $item_list->num_item_types; ++$j) {
                /**
                 * If ThisWorkbook.Worksheets("1.3.Item-Item Compatibility").Cells(k, 3) = "No" Then
                 *      .item_to_item(i, j) = False
                 *      .item_to_item(j, i) = False
                 * End If
                 * k = k + 1
                 *
                 * I see the reason for the hardcoded 3 now. The item compat worksheet starts at row 3.
                 */
                if ($item_item_compatibilities[$i][$j] == false) {
                    $cL->item_to_item[$i][$j] = false;
                }
                if ($item_item_compatibilities[$j][$i] == false) {
                    $cL->item_to_item[$j][$i] = false;
                }

                $k = $k + 1;
                /*
                 * TODO: make sure you get this right.
                 * Right now you're assuming that it's a two dimensional associative array.
                 * */
            }
        }
    }

    /**
     * If instance.container_item_compatibility_worksheet = True Then
     */
    if ($instance->container_item_compatibility_worksheet == true) {
        /**
         * ReDim .container_to_item(1 To container_list.num_container_types, 1 To item_list.num_item_types)
         *
         * Just redimensions the $cL->container_to_item array to a two dimensional array.
         */

        /**
         * For i = 1 To container_list.num_container_types
         */
        for ($i = 1; $i <= $container_list->num_container_types; ++$i) {
            /**
             * For j = 1 To item_list.num_item_types
             */
            for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                /**
                 * .container_to_item(i, j) = True
                 */
                $cL->container_to_item[$i][$j] = true;
            }
        }

        /**
         * k = 3
         */
        $k = 3;

        /**
         * For i = 1 To container_list.num_container_types
         */
        for ($i = 1; $i <= $container_list->num_container_types; ++$i) {
            /**
             * For j = 1 To item_list.num_item_types
             */
            for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                /**
                 * If ThisWorkbook.Worksheets("2.3.Container-ItemCompatibility").Cells(k, 3) = "No" Then
                 *      .container_to_item(i, j) = False
                 * End If
                 *
                 * k = k + 1
                 */
                if ($container_item_compatibilities[$i][$j] == false) {
                    $cL->container_to_item[$i][$j] = false;
                }

                $k = $k + 1;
                /*
                 * TODO: find out if this interpretation is correct.
                 * */
            }
        }
    }
}

/**
 * Private Sub InitializeSolution(solution As solution_data)
 *
 * @param solution_data $solution
 * @param container_list_data $container_list
 * @param item_list_data $item_list
 */
function InitializeSolution(solution_data $solution, container_list_data $container_list, item_list_data $item_list)
{
    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim k As Long
     * @var integer
     */
    $k = 0;

    /**
     * Dim l As Long
     * @var integer
     */
    $l = 0;

    /**
     * With solution
     */
    $s = $solution;

    /**
     * .feasible = False
     */
    $s->feasible = false;

    /**
     * .net_profit = 0
     */
    $s->net_profit = 0;

    /**
     * .total_volume = 0
     */
    $s->total_volume = 0;

    /**
     * .total_weight = 0
     */
    $s->total_weight = 0;

    /**
     * .total_dispersion = 0
     */
    $s->total_dispersion = 0;

    /**
     * .total_distance = 0
     */
    $s->total_distance = 0;

    /**
     * .total_x_moment = 0
     */
    $s->total_x_moment = 0;

    /**
     * .total_yz_moment = 0
     */
    $s->total_yz_moment = 0;

    /**
     * .num_containers = 0
     */
    $s->num_containers = 0;

    /**
     * For i = 1 To container_list.num_container_types
     */
    for ($i = 1; $i <= $container_list->num_container_types; ++$i) {
        /**
         * If container_list.container_types(i).mandatory >= 0 Then
         */
        if ($container_list->container_types[$i]->mandatory >= 0) {
            /**
             * .num_containers = .num_containers + container_list.container_types(i).number_available
             */
            $s->num_containers = $s->num_containers + $container_list->container_types[$i]->number_available;
        }
    }

    /**
     * ReDim .rotation_order(1 To item_list.num_item_types, 1 To 6)
     *
     * I'm guessing this is just a resize. The data fill will soon follow.
     */

    /**
     * For i = 1 To item_list.num_item_types
     */
    for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
        /**
         * For j = 1 To 6
         */
        for ($j = 1; $j <= 6; ++$j) {
            /**
             * .rotation_order(i, j) = j
             */
            $s->rotation_order[$i][$j] = $j;
        }
    }

    /**
     * ReDim .item_type_order(1 To item_list.num_item_types)
     *
     * another resize
     */

    /**
     * For i = 1 To item_list.num_item_types
     */
    for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
        /**
         * .item_type_order(i) = i
         */
        $s->item_type_order[$i] = $i;
    }

    /**
     * ReDim .container(1 To .num_containers)
     *
     * For i = 1 To .num_containers
     *      ReDim .container(i).items(1 To item_list.total_number_of_items)
     *      ReDim .container(i).addition_points(1 To 3 * item_list.total_number_of_items)
     *      ReDim .container(i).repack_item_count(1 To item_list.total_number_of_items)
     * Next i
     *
     * ReDim .unpacked_item_count(1 To item_list.num_item_types)
     *
     * Just a bunch of different resizes.
     */

    /**
     * l = 1
     */
    $l = 1;

    /**
     * For i = 1 To container_list.num_container_types
     */
    for ($i = 1; $i <= $container_list->num_container_types; ++$i) {
        /**
         * If container_list.container_types(i).mandatory >= 0 Then
         */
        if ($container_list->container_types[$i]->mandatory >= 0) {
            /**
             * For j = 1 To container_list.container_types(i).number_available
             */
            for ($j = 1; $j <= $container_list->container_types[$i]->number_available; ++$j) {
                /**
                 * .container(l).width = container_list.container_types(i).width
                 */
                $s->container[$l]->width = $container_list->container_types[$i]->width;

                /**
                 * .container(l).height = container_list.container_types(i).height
                 */
                $s->container[$l]->height = $container_list->container_types[$i]->height;

                /**
                 * .container(l).length = container_list.container_types(i).length
                 */
                $s->container[$l]->length = $container_list->container_types[$i]->length;

                /**
                 * .container(l).volume_capacity = container_list.container_types(i).volume_capacity
                 */
                $s->container[$l]->volume_capacity = $container_list->container_types[$i]->volume_capacity;

                /**
                 * .container(l).weight_capacity = container_list.container_types(i).weight_capacity
                 */
                $s->container[$l]->weight_capacity = $container_list->container_types[$i]->weight_capacity;

                /**
                 * .container(l).cost = container_list.container_types(i).cost
                 */
                $s->container[$l]->cost = $container_list->container_types[$i]->cost;

                /**
                 * .container(l).mandatory = container_list.container_types(i).mandatory
                 */
                $s->container[$l]->mandatory = $container_list->container_types[$i]->mandatory;

                /**
                 * .container(l).type_id = i
                 */
                $s->container[$l]->type_id = $i;

                /**
                 * .container(l).volume_packed = 0
                 */
                $s->container[$l]->volume_packed = 0;

                /**
                 * .container(l).weight_packed = 0
                 */
                $s->container[$l]->weight_packed = 0;

                /**
                 * .container(l).item_cnt = 0
                 */
                $s->container[$l]->item_cnt = 0;

                /**
                 * .container(l).addition_point_count = 1
                 */
                $s->container[$l]->addition_point_count = 1;

                /**
                 * For k = 1 To item_list.total_number_of_items
                 */
                for ($k = 1; $k <= $item_list->total_number_of_items; ++$k) {
                    /**
                     * .container(l).items(k).item_type = 0
                     */
                    $s->container[$l]->items[$k]->item_type = 0;

                    /**
                     * .container(l).addition_points(k).origin_x = 0
                     */
                    $s->container[$l]->addition_points[$k]->origin_x = 0;

                    /**
                     * .container(l).addition_points(k).origin_y = 0
                     */
                    $s->container[$l]->addition_points[$k]->origin_y = 0;

                    /**
                     * .container(l).addition_points(k).origin_z = 0
                     */
                    $s->container[$l]->addition_points[$k]->origin_z = 0;

                    /**
                     * .container(l).addition_points(k).next_to_item_type = 0
                     */
                    $s->container[$l]->addition_points[$k]->next_to_item_type = 0;
                }

                /**
                 * For k = 1 To item_list.total_number_of_items
                 */
                for ($k = 1; $k <= $item_list->total_number_of_items; ++$k) {
                    /**
                     * .container(l).repack_item_count(k) = 0
                     */
                    $s->container[$l]->repack_item_count[$k] = 0;
                }

                /**
                 * l = l + 1
                 */
                $l = $l + 1;
            }
        }
    }

    /**
     * For i = 1 To item_list.num_item_types
     */
    for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
        /**
         * .unpacked_item_count(i) = item_list.item_types(i).number_requested
         */
        $s->unpacked_item_count[$i] = $item_list->item_types[$i]->number_requested;
    }
}

/**
 * Private Sub GetInstanceData()
 *
 * @param instance_data $instance
 * @param bool $front_side_support
 * @param bool $item_item_compatibility_worksheet
 * @param bool $container_item_compatibility_worksheet
 */
function GetInstanceData(instance_data $instance, bool $front_side_support, bool $item_item_compatibility_worksheet, bool $container_item_compatibility_worksheet)
{
    /**
     * If ThisWorkbook.Worksheets("CLP SOlver Console").Cells(6, 3).Value = "Yes" Then
     *      instance.front_side_support = True
     * Else
     *      instance.front_side_support = False
     * End If
     *
     * Just get the value of the $front_side_support (from the probable Request)
     * then set it on the $instance.
     * Just have patience with the sheer amount of redundancy...
     */
    if ($front_side_support == true) {
        $instance->front_side_support = true;
    } else {
        $instance->front_side_support = false;
    }

    /**
     * If CheckWorksheetExistence("1.3.Item-Item Compatibility") = True Then
     *      instance.item_item_compatibility_worksheet = True
     * Else
     *      instance.item_item_compatibility_worksheet = False
     * End If
     *
     * This checks if the optional "1.3.Item-Item Compatibility" exists and sets the corresponding
     * prop on the instance.
     * In our case, we're passing in the boolean $item_item_compatibility_worksheet
     */
    if ($item_item_compatibility_worksheet == true) {
        $instance->item_item_compatibility_worksheet = true;
    } else {
        $instance->item_item_compatibility_worksheet = false;
    }

    /**
     * If CheckWorksheetExistence("2.3.Container-ItemCompatibility") = True Then
     *      instance.container_item_compatibility_worksheet = True
     * Else
     *      instance.container_item_compatibility_worksheet = False
     * End If
     *
     * Also checks for the existence of the "2.3.Container-ItemCompatibility" worksheet, then
     * sets the appropriate flag on the instance.
     *
     * We'll just receive a boolean $container_item_compatibility_worksheet to handle this translation.
     */
    if ($container_item_compatibility_worksheet == true) {
        $instance->container_item_compatibility_worksheet = true;
    } else {
        $instance->container_item_compatibility_worksheet = false;
    }
}

/**
 * Private Sub WriteSolution(solution As solution_data)
 *
 * @param solution_data $solution
 * @param container_list_data $container_list
 * @param item_list_data $item_list
 */
function WriteSolution(solution_data $solution, container_list_data $container_list, item_list_data $item_list)
{
    /**
     * Application.ScreenUpdating = False
     * Application.Calculation = xlCalculationManual
     *
     * I'm guessing this controls something in MS Excel.
     */

    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim k As Long
     * @var integer
     */
    $k = 0;

    /**
     * Dim container_index As Long
     * @var integer
     */
    $container_index = 0;

    /**
     * Dim swap_container As container_data
     * @var container_data
     */
    $swap_container = new container_data();

    /*
     * 'sort the containers
     * */

    /**
     * For i = 1 To solution.num_containers
     */
    for ($i = 1; $i <= $solution->num_containers; ++$i) {
        /**
         * For j = solution.num_containers To 2 Step -1
         */
        for ($j = $solution->num_containers; $j >= 2; --$j) {
            /**
             * If (solution.container(j).type_id < solution.container(j - 1).type_id) Or _
             *      ((solution.container(j).type_id = solution.container(j - 1).type_id) And (solution.container(j).volume_packed > solution.container(j - 1).volume_packed)) Then
             */
            if (
                $solution->container[$j]->type_id < $solution->container[$j - 1]->type_id ||
                (
                    $solution->container[$j]->type_id == $solution->container[$j - 1]->type_id &&
                    $solution->container[$j]->volume_packed > $solution->container[$j - 1]->volume_packed
                )
            ) {
                /**
                 * swap_container = solution.container(j)
                 */
                $swap_container = $solution->container[$j];

                /**
                 * solution.container(j) = solution.container(j - 1)
                 */
                $solution->container[$j] = $solution->container[$j - 1];

                /**
                 * solution.container(j - 1) = swap_container
                 */
                $solution->container[$j - 1] = $swap_container;
            }
        }
    }

    /**
     * ThisWorkbook.Worksheets("3.Solution").Activate
     *
     * Bring the "3.Solution" worksheet into focus (?)
     */

    $warning = "";
    /*
     * The following code deals with some warning,
     * so I set this little variable here. We'll
     * see how it will work its way into a future
     * app.
     * */

    /**
     * If solution.feasible = False Then
     */
    if ($solution->feasible == false) {
        /**
         * Cells(2, 1) = "Warning: Last solution returned by the solver does not satisfy all constraints."
         * Range(Cells(2, 1), Cells(2, 10)).Interior.ColorIndex = 45
         *
         * Umm, I guess this throws a warning... and changes some colors
         */
        /*
         * TODO: incorporate this as a warning in a future app.
         * */
        $warning = "Warning: Last solution returned by the solver does not satisfy all constraints.";
    } else {
        /**
         * Cells(2, 1) = vbNullString
         * Range(Cells(2, 1), Cells(2, 10)).Interior.Pattern = xlNone
         * Range(Cells(2, 1), Cells(2, 10)).Interior.TintAndShade = 0
         * Range(Cells(2, 1), Cells(2, 10)).Interior.PatternTintAndShade = 0
         *
         * This would be some CSS changes on the frontend, in our case.
         */
        $warning = "";
    }

    /**
     * Dim offset As Long
     * @var integer
     */
    $offset = 0;

    /**
     * offset = 0
     *
     * variable already set
     */

    /**
     * container_index = 1
     */
    $container_index = 1;

    /**
     * With solution
     *
     * We'll just shorten $solution to $s
     */
    $s = $solution;

    /**
     * For i = 1 To container_list.num_container_types
     */
    for ($i = 1; $i <= $container_list->num_container_types; ++$i) {
        /**
         * For j = 1 To container_list.container_types(i).number_available
         */
        for ($j = 1; $j <= $container_list->container_types[$i]->number_available; ++$j) {
            /**
             * Range(Cells(6, offset + 2), Cells(5 + 2 * item_list.total_number_of_items, offset + 2)).Value = vbNullString
             * Range(Cells(6, offset + 3), Cells(5 + 2 * item_list.total_number_of_items, offset + 5)).ClearContents
             * Range(Cells(6, offset + 6), Cells(5 + 2 * item_list.total_number_of_items, offset + 6)).Value = vbNullString
             *
             * I guess this clears some cells in the worksheet.
             */

            /**
             * If container_list.container_types(i).mandatory >= 0 Then
             */
            if ($container_list->container_types[$i]->mandatory >= 0) {
                /**
                 * For k = 1 To .container(container_index).item_cnt
                 */
                for ($k = 1; $k <= $s->container[$container_index]->item_cnt; ++$k) {
                    /**
                     * Cells(5 + k, offset + 2).Value = ThisWorkbook.Worksheets("1.Items").Cells(2 + item_list.item_types(.container(container_index).items(k).item_type).id, 2).Value
                     * Cells(5 + k, offset + 3).Value = .container(container_index).items(k).origin_x
                     * Cells(5 + k, offset + 4).Value = .container(container_index).items(k).origin_y
                     * Cells(5 + k, offset + 5).Value = .container(container_index).items(k).origin_z
                     *
                     * This sets some values on the worksheets.
                     */

                    /*
                     * TODO: transpose the following orientation strings to something useful in an app.
                     * */

                    /**
                     * If .container(container_index).items(k).rotation = 1 Then
                     */
                    if ($s->container[$container_index]->items[$k]->rotation == 1) {
                        /**
                         * Cells(5 + k, offset + 6).Value = "xyz"
                         */
                    /**
                     * ElseIf .container(container_index).items(k).rotation = 2 Then
                     */
                    } else if ($s->container[$container_index]->items[$k]->rotation == 2) {
                        /**
                         * Cells(5 + k, offset + 6).Value = "zyx"
                         */
                    /**
                     * ElseIf .container(container_index).items(k).rotation = 3 Then
                     */
                    } else if ($s->container[$container_index]->items[$k]->rotation == 3) {
                        /**
                         * Cells(5 + k, offset + 6).Value = "xzy"
                         */
                    /**
                     * ElseIf .container(container_index).items(k).rotation = 4 Then
                     */
                    } else if ($s->container[$container_index]->items[$k]->rotation == 4) {
                        /**
                         * Cells(5 + k, offset + 6).Value = "yzx"
                         */
                    /**
                     * ElseIf .container(container_index).items(k).rotation = 5 Then
                     */
                    } else if ($s->container[$container_index]->items[$k]->rotation == 5) {
                        /**
                         * Cells(5 + k, offset + 6).Value = "yxz"
                         */
                    /**
                     * ElseIf .container(container_index).items(k).rotation = 6 Then
                     */
                    } else if ($s->container[$container_index]->items[$k]->rotation == 5) {
                        /**
                         * Cells(5 + k, offset + 6).Value = "zxy"
                         */
                    }
                }

                /**
                 * container_index = container_index + 1
                 */
                $container_index = $container_index + 1;
            }
            /**
             * Columns(2 + offset).AutoFit
             *
             * Ummm... maybe this resizes the columns so everything looks clean in the Excel worksheet.
             */
            /**
             * offset = offset + column_offset
             */
            $offset = $offset + column_offset;
        }
    }

    /**
     * Range(Cells(6, offset + 2), Cells(5 + 2 * item_list.total_number_of_items, offset + 2)).Value = vbNullString
     *
     * Writes nothing to a bunch of cells
     */

    /**
     * k = 1
     */
    $k = 1;

    /**
     * For i = 1 To item_list.num_item_types
     */
    for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
        /**
         * For j = 1 To .unpacked_item_count(i)
         */
        for ($j = 1; $j <= $s->unpacked_item_count[$i]; ++$j) {
            /**
             * Cells(5 + k, offset + 2).Value = ThisWorkbook.Worksheets("1.Items").Cells(2 + item_list.item_types(i).id, 2).Value
             *
             * Moves some values from one place to another (?)
             */
            /**
             * k = k + 1
             */
            $k = $k + 1;
        }
    }

    /**
     * End With
     *
     * Application.ScreenUpdating = True
     * Application.Calculation = xlCalculationAutomatic
     */
}

/**
 * Sub ReadSolution(solution As solution_data)
 *
 *
 */
function ReadSolution(solution_data $solution, container_list_data $container_list, item_list_data $item_list)
{
    /**
     * Application.ScreenUpdating = False
     * Application.Calculation = xlCalculationManual
     *
     * Excel stuff
     */

    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim k As Long
     * @var integer
     */
    $k = 0;

    /**
     * Dim l As Long
     * @var integer
     */
    $l = 0;

    /**
     * Dim container_index As Long
     * @var integer
     */
    $container_index = 0;

    /**
     * Dim item_type_index As Long
     * @var integer
     */
    $item_type_index = 0;

    /**
     * Dim offset As Long
     * @var integer
     */
    $offset = 0;

    /**
     * offset = 0
     *
     * already defined
     */

    /**
     * container_index = 1
     */
    $container_index = 1;

    /**
     * With solution
     * $s for $solution
     */
    $s = $solution;

    /**
     * For i = 1 To container_list.num_container_types
     */
    for ($i = 1; $i <= $container_list->num_container_types; ++$i) {
        /**
         * For j = 1 To container_list.container_types(i).number_available
         */
        for ($j = 1; $j <= $container_list->container_types[$i]->number_available; ++$j) {
            /**
             * If container_list.container_types(i).mandatory >= 0 Then
             */
            if ($container_list->container_types[$i]->mandatory >= 0) {
                /**
                 * With .container(container_index)
                 *
                 * we'll call this $cC, short for $currentContainer;
                 */
                $cC = $s->container[$container_index];

                /**
                 * l = Cells(4, offset + 7).Value
                 *
                 * Umm...context please?
                 * I guess it reads from the Solution worksheet. At row 4, column 0 + 7.
                 * I just need to know when this happens.
                 * The formula for these cells is a COUNT, which determines how many items
                 * are in the container.
                 */

                /**
                 * Since this reads from a spreadsheet and we currently have no alternative,
                 * we'll just let this and the following code commented.
                 *
                 *  For k = 1 To l
                 *      If IsNumeric(Cells(5 + k, offset + 7).Value) = True Then
                 *          .item_cnt = .item_cnt + 1
                 *          item_type_index = Cells(5 + k, offset + 7).Value
                 *
                 *          solution.unpacked_item_count(item_type_index) = solution.unpacked_item_count(item_type_index) - 1
                 *
                 *          .items(.item_cnt).item_type = item_type_index
                 *          .items(.item_cnt).origin_x = Cells(5 + k, offset + 3).Value
                 *          .items(.item_cnt).origin_y = Cells(5 + k, offset + 4).Value
                 *          .items(.item_cnt).origin_z = Cells(5 + k, offset + 5).Value
                 *
                 *          If ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 6).Value = "xyz" Then
                 *              .items(.item_cnt).rotation = 1
                 *              .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).width
                 *              .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).height
                 *              .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).length
                 *          ElseIf ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 6).Value = "zyx" Then
                 *              .items(.item_cnt).rotation = 2
                 *              .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).length
                 *              .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).height
                 *              .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).width
                 *          ElseIf ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 6).Value = "xzy" Then
                 *              .items(.item_cnt).rotation = 3
                 *              .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).width
                 *              .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).length
                 *              .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).height
                 *          ElseIf ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 6).Value = "yzx" Then
                 *              .items(.item_cnt).rotation = 4
                 *              .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).height
                 *              .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).length
                 *              .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).width
                 *          ElseIf ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 6).Value = "yxz" Then
                 *              .items(.item_cnt).rotation = 5
                 *              .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).height
                 *              .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).width
                 *              .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).length
                 *          ElseIf ThisWorkbook.Worksheets("3.Solution").Cells(5 + k, offset + 6).Value = "zxy" Then
                 *              .items(.item_cnt).rotation = 6
                 *              .items(.item_cnt).opposite_x = .items(.item_cnt).origin_x + item_list.item_types(item_type_index).length
                 *              .items(.item_cnt).opposite_y = .items(.item_cnt).origin_y + item_list.item_types(item_type_index).width
                 *              .items(.item_cnt).opposite_z = .items(.item_cnt).origin_z + item_list.item_types(item_type_index).height
                 *          End If
                 *
                 *          .volume_packed = .volume_packed + item_list.item_types(item_type_index).volume
                 *          .weight_packed = .weight_packed + item_list.item_types(item_type_index).weight
                 *
                 *          If .item_cnt = 1 Then
                 *              solution.net_profit = solution.net_profit + item_list.item_types(item_type_index).profit - .cost
                 *          Else
                 *              solution.net_profit = solution.net_profit + item_list.item_types(item_type_index).profit
                 *          End If
                 *
                 *      End If
                 * Next k
                 */

                /*
                 * TODO: remember to find a solution for this. This looks like two-way binding (reading to and from the spreadsheet), a modern idea in web dev.
                 * */

                /**
                 * End With
                 */

                /**
                 * container_index = container_index + 1
                 */
                $container_index = $container_index + 1;
            }
            $offset = $offset + column_offset;
        }
    }
    /**
     * End With
     */
    /**
     * Application.ScreenUpdating = True
     * Application.Calculation = xlCalculationAutomatic
     */
}

/**
 * Sub CLP_Solver()
 *
 */
function CLP_Solver()
{
    /**
     * Application.ScreenUpdating = False
     * Application.Calculation = xlCalculationManual
     *
     * Excel stuff.
     */

    /**
     * Dim WorksheetExists As Boolean
     * @var bool
     */
    $WorksheetExists = false;
    /*
     * This isn't really necessary, since we're dealing with worksheets.
     * */

    /**
     * Dim reply As Integer
     * @var integer
     */
    $reply = 0;

    /**
     * WorksheetExists = CheckWorksheetExistence("1.Items") And CheckWorksheetExistence("2.Containers") And CheckWorksheetExistence("3.Solution")
     */
    $WorksheetExists = CheckWorksheetExistence("1.Items") && CheckWorksheetExistence("2.Containers") && CheckWorksheetExistence("3.Solution");

    /**
     * If WorksheetExists = False Then
     */
    if ($WorksheetExists == false) {
        /**
         * MsgBox "Worksheets 1.Items, 2.Containers, and 3.Solution must exist for the CLP Spreadsheet Solver to function."
         * Application.ScreenUpdating = True
         * Application.Calculation = xlCalculationAutomatic
         * Exit Sub
         */
        return null;
    } else {
        /**
         * reply = MsgBox("This will take " & ThisWorkbook.Worksheets("CLP Solver Console").Cells(14, 3).Value & " seconds. Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
         * If reply = vbNo Then
         *      Application.ScreenUpdating = True
         *      Application.Calculation = xlCalculationAutomatic
         *      Exit Sub
         * End If
         */
    }

    /**
     * Application.EnableCancelKey = xlErrorHandler
     * On Error GoTo CLP_Solver_Finish
     *
     * Hmm, should try to at least understand this. Maybe later...
     */

    /*
     * 'Allocate memory and get the data
     * */
    /**
     * Call GetSolverOptions
     *
     * We'll need to pass in some params for our translated functions.
     */
    $solver_options = new solver_option_data();
    $item_sort_criterion = "";
    $wall_building = true;
    $CPU_time_limit = 60;
    GetSolverOptions($solver_options, $item_sort_criterion, $wall_building, $CPU_time_limit);

    /**
     * Call GetItemData
     */
    $item_list = new item_list_data();
    $num_item_types = 1;
    $items = array();
    $solver_options = new solver_option_data();
    GetItemData($item_list, $num_item_types, $items, $solver_options);

    /**
     * Call GetContainerData
     *
     */
    $container_list = new container_list_data();
    $num_container_types = 1;
    $containers = array();
    GetContainerData($container_list, $num_container_types, $containers);

    /**
     * Call GetInstanceData
     */
    $instance = new instance_data();
    $front_side_support = false;
    $item_item_compatibility_worksheet = false;
    $container_item_compatibility_worksheet = false;
    GetInstanceData($instance, $front_side_support, $item_item_compatibility_worksheet, $container_item_compatibility_worksheet);

    /**
     * Call GetCompatibilityData
     */
    $compatibility_list = new compatibility_data();
    $instance = new instance_data();
    $item_list = new item_list_data();
    $container_list = new container_list_data();
    $item_item_compatibilities = array();
    $container_item_compatibilities = array();
    GetCompatibilityData($compatibility_list, $instance, $item_list, $container_list, $item_item_compatibilities, $container_item_compatibilities);

    /**
     * Call SortItems
     */
    SortItems();

    /**
     * Dim incumbent As solution_data
     * @var solution_data
     */
    $incumbent = new solution_data();

    /**
     * Call InitializeSolution(incumbent)
     */
    $container_list = new container_list_data();
    $item_list = new item_list_data();
    InitializeSolution($incumbent, $container_list, $item_list);

    /**
     * Dim best_known As solution_data
     * @var solution_data
     */
    $best_known = new solution_data();

    /**
     * Call InitializeSolution(best_known)
     */
    $container_list = new container_list_data();
    $item_list = new item_list_data();
    InitializeSolution($best_known, $container_list, $item_list);

    /**
     * best_known = incumbent
     */
    $best_known = $incumbent;

    /**
     * Dim iteration As Long
     * @var integer
     */
    $iteration = 0;

    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim k As Long
     * @var integer
     */
    $k = 0;

    /**
     * Dim l As Long
     * @var integer
     */
    $l = 0;

    /**
     * Dim nonempty_container_cnt As Long
     * @var integer
     */
    $nonempty_container_cnt = 0;

    /**
     * Dim container_id As Long
     * @var integer
     */
    $container_id = 0;

    /**
     * Dim start_time As Double
     * @var float
     */
    $start_time = 0.0;

    /**
     * Dim end_time As Double
     * @var float
     */
    $end_time = 0.0;

    /**
     * Dim continue_flag As Boolean
     * @var boolean
     */
    $continue_flag = true;

    /**
     * Dim sort_criterion As Double
     * @var float
     */
    $sort_criterion = 0.0;

    /**
     * Dim selected_rotation As Double
     * @var float
     */
    $selected_rotation = 0.0;

    /*
     * 'infeasibility check
     * */

    /**
     * Dim infeasibility_count As Long
     * @var integer
     */
    $infeasibility_count = 0;

    /**
     * Dim infeasibility_string As Long
     * @var string
     */
    $infeasibility_string = 0;

    /**
     * Call FeasibilityCheckData(infeasibility_count, infeasibility_string)
     */
    FeasibilityCheckData($infeasibility_count, $infeasibility_string);

    /**
     * If infeasibility_count > 0 Then
     */
    if ($infeasibility_count > 0) {
        /**
         * reply = MsgBox("Infeasibilities detected." & Chr(13) & infeasibility_string & "Do you want to continue?", vbYesNo, "CLP Spreadsheet Solver")
         * If reply = vbNo Then
         *      Application.ScreenUpdating = True
         *      Application.Calculation = xlCalculationAutomatic
         *      Exit Sub
         * End If
         */
    }

    /**
     * start_time = Timer
     * end_time = Timer
     *
     * We don't really need these
     */

    /*
     * 'constructive phase
     * */

    /**
     * Application.ScreenUpdating = True
     * Application.StatusBar = "Constructive phase..."
     * Application.ScreenUpdating = False
     */

    /**
     * Call SortContainers(incumbent, 0)
     */
    SortContainers($incumbent, 0);

    /**
     * For i = 1 To incumbent.num_containers
     */
    for ($i = 1; $i <= $incumbent->num_containers; ++$i) {
        /*
         * 'sort the rotation order for this container
         * */
        /**
         * For j = 1 To item_list.num_item_types
         */
        for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
            /**
             * sort_criterion = 0
             */
            $sort_criterion = 0;

            /**
             * selected_rotation = 0
             */
            $selected_rotation = 0;

            /**
             * If sort_criterion < (Int(incumbent.container(i).width / item_list.item_types(j).width) * item_list.item_types(j).width) * (Int(incumbent.container(i).height / item_list.item_types(j).height) * item_list.item_types(j).height) Then
             *      sort_criterion = (Int(incumbent.container(i).width / item_list.item_types(j).width) * item_list.item_types(j).width) * (Int(incumbent.container(i).height / item_list.item_types(j).height) * item_list.item_types(j).height)
             *      selected_rotation = 1
             * End If
             */
            if (
                $sort_criterion < (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->width) * $item_list->item_types[$j]->width) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->height) * $item_list->item_types[$j]->height)
            ) {
                $sort_criterion = (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->width) * $item_list->item_types[$j]->width) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->height) * $item_list->item_types[$j]->height);
                $selected_rotation = 1;
            }

            /**
             * If sort_criterion < (Int(incumbent.container(i).width / item_list.item_types(j).length) * item_list.item_types(j).length) * (Int(incumbent.container(i).height / item_list.item_types(j).height) * item_list.item_types(j).height) Then
             *      sort_criterion = (Int(incumbent.container(i).width / item_list.item_types(j).length) * item_list.item_types(j).length) * (Int(incumbent.container(i).height / item_list.item_types(j).height) * item_list.item_types(j).height)
             *      selected_rotation = 2
             * End If
             */
            if (
                $sort_criterion < (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->length) * $item_list->item_types[$j]->length) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->height) * $item_list->item_types[$j]->height)
            ) {
                $sort_criterion = (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->length) * $item_list->item_types[$j]->length) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->height) * $item_list->item_types[$j]->height);
                $selected_rotation = 2;
            }

            /**
             * If (item_list.item_types(j).xy_rotatable = True) And (sort_criterion < (Int(incumbent.container(i).width / item_list.item_types(j).width) * item_list.item_types(j).width) * (Int(incumbent.container(i).height / item_list.item_types(j).length) * item_list.item_types(j).length)) Then
             *      sort_criterion = (Int(incumbent.container(i).width / item_list.item_types(j).width) * item_list.item_types(j).width) * (Int(incumbent.container(i).height / item_list.item_types(j).length) * item_list.item_types(j).length)
             *      selected_rotation = 3
             * End If
             */
            if (
                $item_list->item_types[$j]->xy_rotatable == true &&
                $sort_criterion < (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->width) * $item_list->item_types[$j]->width) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->length) * $item_list->item_types[$j]->length)
            ) {
                $sort_criterion = (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->width) * $item_list->item_types[$j]->width) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->length) * $item_list->item_types[$j]->length);
                $selected_rotation = 3;
            }

            /**
             * If (item_list.item_types(j).xy_rotatable = True) And (sort_criterion < (Int(incumbent.container(i).width / item_list.item_types(j).height) * item_list.item_types(j).height) * (Int(incumbent.container(i).height / item_list.item_types(j).length) * item_list.item_types(j).length)) Then
             *      sort_criterion = sort_criterion < (Int(incumbent.container(i).width / item_list.item_types(j).height) * item_list.item_types(j).height) * (Int(incumbent.container(i).height / item_list.item_types(j).length) * item_list.item_types(j).length)
             *      selected_rotation = 4
             * End If
             */
            if (
                $item_list->item_types[$j]->xy_rotatable == true &&
                $sort_criterion < (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->height) * $item_list->item_types[$j]->height) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->length) * $item_list->item_types[$j]->length)
            ) {
                $sort_criterion = (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->height) * $item_list->item_types[$j]->height) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->length) * $item_list->item_types[$j]->length);
                $selected_rotation = 4;
            }

            /**
             * If (item_list.item_types(j).yz_rotatable = True) And (sort_criterion < (Int(incumbent.container(i).width / item_list.item_types(j).height) * item_list.item_types(j).height) * (Int(incumbent.container(i).height / item_list.item_types(j).width) * item_list.item_types(j).width)) Then
             *      sort_criterion = (Int(incumbent.container(i).width / item_list.item_types(j).height) * item_list.item_types(j).height) * (Int(incumbent.container(i).height / item_list.item_types(j).width) * item_list.item_types(j).width)
             *      selected_rotation = 5
             * End If
             */
            if (
                $item_list->item_types[$j]->yz_rotatable == true &&
                $sort_criterion < (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->height) * $item_list->item_types[$j]->height) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->width) * $item_list->item_types[$j]->width)
            ) {
                $sort_criterion = (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->height) * $item_list->item_types[$j]->height) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->width) * $item_list->item_types[$j]->width);
                $selected_rotation = 5;
            }

            /**
             * If (item_list.item_types(j).yz_rotatable = True) And (sort_criterion < (Int(incumbent.container(i).width / item_list.item_types(j).length) * item_list.item_types(j).length) * (Int(incumbent.container(i).height / item_list.item_types(j).width) * item_list.item_types(j).width)) Then
             *      sort_criterion = (Int(incumbent.container(i).width / item_list.item_types(j).length) * item_list.item_types(j).length) * (Int(incumbent.container(i).height / item_list.item_types(j).width) * item_list.item_types(j).width)
             *      selected_rotation = 6
             * End If
             */
            if (
                $item_list->item_types[$j]->yz_rotatable == true &&
                $sort_criterion < (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->length) * $item_list->item_types[$j]->length) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->width) * $item_list->item_types[$j]->width)
            ) {
                $sort_criterion = (intval($incumbent->container[$i]->width / $item_list->item_types[$j]->length) * $item_list->item_types[$j]->length) * (intval($incumbent->container[$i]->height / $item_list->item_types[$j]->width) * $item_list->item_types[$j]->width);
                $selected_rotation = 6;
            }

            /**
             * TODO: double and triple check this. really long lines.
             */

            /**
             * If selected_rotation = 0 Then
             *      selected_rotation = 1
             * End If
             */
            if ($selected_rotation == 0) {
                $selected_rotation = 1;
            }

            /**
             * incumbent.rotation_order(j, 1) = selected_rotation
             */
            $incumbent->rotation_order[$j][1] = $selected_rotation;

            /**
             * incumbent.rotation_order(j, selected_rotation) = 1
             */
            $incumbent->rotation_order[$j][$selected_rotation] = 1;
        }

        /**
         * For j = 1 To item_list.num_item_types
         */
        for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
            /**
             * continue_flag = True
             */
            $continue_flag = true;
            /**
             * Do While (incumbent.unpacked_item_count(incumbent.item_type_order(j)) > 0) And (continue_flag = True)
             *      continue_flag = AddItemToContainer(incumbent, i, incumbent.item_type_order(j), 1, False)
             * Loop
             */
            do {
                $continue_flag = AddItemToContainer($incumbent, $i, $incumbent->item_type_order[$j], 1, false, $instance, $compatibility_list, $item_list, $solver_options);
            } while ($incumbent->unpacked_item_count[$incumbent->item_type_order[$j]] > 0 && $continue_flag == true);
        }

        /**
         * incumbent.feasible = True
         */
        $incumbent->feasible = true;

        /**
         * For j = 1 To item_list.num_item_types
         */
        for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
            /**
             * If (incumbent.unpacked_item_count(j) > 0) And (item_list.item_types(j).mandatory = 1) Then
             */
            if ($incumbent->unpacked_item_count[$j] > 0 && $item_list->item_types[$j]->mandatory == 1) {
                /**
                 * incumbent.feasible = False
                 */
                $incumbent->feasible = false;
                /**
                 * Exit For
                 */
                break;
            }
        }

        /**
         * Call CalculateDispersion(incumbent)
         */
        CalculateDispersion($incumbent);

        /**
         * If ((incumbent.feasible = True) And (best_known.feasible = False)) Or _
         *      ((incumbent.feasible = False) And (best_known.feasible = False) And (incumbent.total_volume > best_known.total_volume + epsilon)) Or _
         *      ((incumbent.feasible = True) And (best_known.feasible = True) And (incumbent.net_profit > best_known.net_profit + epsilon)) Or _
         *      ((incumbent.feasible = True) And (best_known.feasible = True) And (incumbent.net_profit > best_known.net_profit - epsilon) And (incumbent.total_volume > best_known.total_volume + epsilon)) Then
         */
        if (
            ($incumbent->feasible == true && $best_known->feasible == false) ||
            ($incumbent->feasible == false && $best_known->feasible == false && $incumbent->total_volume > $best_known->total_volume + epsilon) ||
            ($incumbent->feasible == true && $best_known->feasible == true && $incumbent->net_profit > $best_known->net_profit + epsilon) ||
            ($incumbent->feasible == true && $best_known->feasible == true && $incumbent->net_profit > $best_known->net_profit - epsilon && $incumbent->total_volume > $best_known->total_volume + epsilon)
        ) {
            /**
             * best_known = incumbent
             */
            $best_known = $incumbent;
        }
    }

    /*
     * 'GoTo CLP_Solver_Finish
     *
     * 'end_time = Timer
     * 'MsgBox "Constructive phase result: " & best_known.net_profit & " time: " & end_time - start_time
     *
     * 'improvement phase
     * */
    /**
     * iteration = 0
     */
    $iteration = 0;

    /**
     * Do
     *
     * Well, this is different. There are while and do while loops in php, I wonder if
     * I treated the problem correctly so far.
     */
    do {
        /**
         * If iteration Mod 100 = 0 Then
         */
        if ($iteration % 100 == 0) {
            /**
             * Application.ScreenUpdating = True
             * I'll ignore this.
             */
            /**
             * If best_known.feasible = True Then
             */
            if ($best_known->feasible == true) {
                /**
                 * Application.StatusBar = "Starting iteration " & iteration & ". Best net profit found so far: " & best_known.net_profit & " Dispersion: " & best_known.total_dispersion
                 * oh, so this just writes to the excel statusbar something.
                 * ignoring.
                 */
            /**
             * Else
             */
            } else {
                /**
                 * Application.StatusBar = "Starting iteration " & iteration & ". Best net profit found so far: N/A" & " Dispersion: " & best_known.total_dispersion
                 */
            }
            /**
             * Application.ScreenUpdating = False
             * DoEvents
             */
        }

        /**
         * If Rnd < 0.5 Then ' < ((end_time - start_time) / solver_options.CPU_time_limit) ^ 2 Then
         */
        if (random() < 0.5) {
            /**
             * incumbent = best_known
             */
            $incumbent = $best_known;
        }

        /**
         * With incumbent
         */
        $inc = $incumbent;

        /**
         * For i = 1 To .num_containers
         */
        for ($i = 1; $i <= $inc->num_containers; ++$i) {
            /**
             * Call PerturbSolution(incumbent, i, 1 - ((end_time - start_time) / solver_options.CPU_time_limit))
             */
            PerturbSolution($incumbent, $i, 1-(($end_time - $start_time) / $solver_options->CPU_time_limit), $item_list);
        }
        /**
         * Call PerturbRotationAndOrderOfItems(incumbent)
         */
        PerturbRotationAndOrderOfItems($incumbent, $item_list);

        /**
         * End With
         */

        /**
         * Call SortContainers(incumbent, 0.2)
         */
        SortContainers($incumbent, 0.2);

        /**
         * With incumbent
         */
        $inc = $incumbent;

        /**
         * For i = 1 To .num_containers
         */
        for ($i = 1; $i <= $inc->num_containers; ++$i) {
            /**
             * For j = 1 To item_list.num_item_types
             */
            for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                /**
                 * continue_flag = True
                 */
                $continue_flag = true;

                /**
                 * Do While (.unpacked_item_count(.item_type_order(j)) > 0) And (continue_flag = True)
                 *      continue_flag = AddItemToContainer(incumbent, i, .item_type_order(j), 1, False)
                 *      'DoEvents
                 * Loop
                 */
                do {
                    $continue_flag = AddItemToContainer($incumbent, $i, $inc->item_type_order[$j], 1, false, $instance, $compatibility_list, $item_list, $solver_options);
                } while ($inc->unpacked_item_count[$inc->item_type_order[$j]] > 0 && $continue_flag == true);
            }
            /**
             * .feasible = True
             */
            $inc->feasible = true;

            /**
             * For j = 1 To item_list.num_item_types
             */
            for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                /**
                 * If (.unpacked_item_count(j) > 0) And (item_list.item_types(j).mandatory = 1) Then
                 */
                if ($inc->unpacked_item_count[$j] > 0 && $item_list->item_types[$j]->mandatory == 1) {
                    /**
                     * .feasible = False
                     */
                    $inc->feasible = false;

                    /**
                     * Exit For
                     */
                    break;
                }
            }

            /**
             * If .feasible = True Then
             */
            if ($inc->feasible == true) {
                /**
                 * Call CalculateDispersion(incumbent)
                 */
                CalculateDispersion($incumbent);
            }

            /**
             * If ((.feasible = True) And (best_known.feasible = False)) Or _
             *    ((.feasible = False) And (best_known.feasible = False) And (.total_volume > best_known.total_volume + epsilon)) Or _
             *    ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit + epsilon)) Or _
             *    ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_volume > best_known.total_volume + epsilon)) Or _
             *    ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_volume > best_known.total_volume - epsilon) And (.total_dispersion < best_known.total_dispersion - epsilon)) Then
             *
             *     best_known = incumbent
             *
             * End If
             */

            /*
             * TODO: triple check this -> really long lines of conditions
             * */
            if (
                ($inc->feasible == true && $best_known->feasible == false) ||
                ($inc->feasible == false && $best_known->feasible == false && $inc->total_volume > $best_known->total_volume + epsilon) ||
                ($inc->feasible == true && $best_known->feasible == true && $inc->net_profit > $best_known->net_profit + epsilon) ||
                ($inc->feasible == true && $best_known->feasible == true && $inc->net_profit > $best_known->net_profit - epsilon && $inc->total_volume > $best_known->total_volume + epsilon) ||
                ($inc->feasible == true && $best_known->feasible == true && $inc->net_profit > $best_known->net_profit - epsilon && $inc->total_volume > $best_known->total_volume - epsilon && $inc->total_dispersion < $best_known->total_dispersion - epsilon)
            ) {
                $best_known = $incumbent;
            }
        }
        /**
         * End With
         *
         * I'm guessing that this is the original with,
         * the other ones were nested, so the scope was much more localised.
         * (This is just my uninformed opinion/guess, I don't know if it's actually true)
         */

        /**
         * iteration = iteration + 1
         */
        $iteration = $iteration + 1;

        /**
         * end_time = Timer
         *
         * Ummm, I don't remember if I actually did anything with the timer.
         * It's not needed in web apps, usually. We'll just have to figure
         * out an exit strategy later on.
         */

        /**
         * If end_time < start_time - 0.01 Then
         *      solver_options.CPU_time_limit = solver_options.CPU_time_limit - (86400 - start_time)
         *      start_time = end_time
         * End If
         *
         * I'll just translate this for the sake of translation.
         */
        if ($end_time < $start_time - 0.01) {
            $solver_options->CPU_time_limit = $solver_options->CPU_time_limit - (86400 - $start_time);
            $start_time = $end_time;
        }

    /**
     * Loop While end_time - start_time < solver_options.CPU_time_limit / 3
     *
     * Again, this condition needs to be tied into the exit strategy.
     */
    /*
     * TODO: figure out an exit strategy that doesn't rely on time.
     * Reasoning: if we're tied into time as an exit strategy, then
     * we are bound by hardware: weaker computers/servers will run less iterations
     * in the given amount of time than better computers/servers.
     * One way we should devise this exit strategy is by tracking improvement somehow.
     * When there's less improvement in the process than a given threshold -> exit.
     * Also, we would need to somehow do a naive precheck if there's even a possibility
     * for improvement. Work on this later.
     * */
    } while ($end_time - $start_time < $solver_options->CPU_time_limit / 3);

    /*
     * ' reorganize now
     * */

    /**
     * nonempty_container_cnt = 0
     */
    $nonempty_container_cnt = 0;

    /**
     * With best_known
     *
     * We'll just call this $bk
     */
    $bk = $best_known;

    /**
     * For i = 1 To .num_containers
     */
    for ($i = 1; $i <= $bk->num_containers; ++$i) {
        /**
         * If .container(i).item_cnt > 0 Then
         */
        if ($bk->container[$i]->item_cnt > 0) {
            /**
             * nonempty_container_cnt = nonempty_container_cnt + 1
             */
            $nonempty_container_cnt = $nonempty_container_cnt + 1;
        }
    }
    /**
     * End With
     */

    /**
     * For container_id = 1 To best_known.num_containers
     */
    for ($container_id = 1; $container_id <= $best_known->num_containers; ++$container_id) {
        /**
         * Call CalculateDistance(best_known, container_id)
         */
        CalculateDistance($best_known, $container_id);

        /**
         * Application.ScreenUpdating = True
         * If best_known.feasible = True Then
         *      Application.StatusBar = "Reorganizing container " & container_id & ". Best net profit found so far: " & best_known.net_profit & " Distance: " & best_known.total_distance
         * Else
         *      Application.StatusBar = "Reorganizing container " & container_id & ". Best net profit found so far: N/A" & " Distance: " & best_known.total_distance
         * End If
         * Application.ScreenUpdating = False
         *
         * This just writes stuff to the status bar in MS Excel.
         */

        /**
         * If best_known.container(container_id).item_cnt > 0 Then
         */
        if ($best_known->container[$container_id]->item_cnt > 0) {
            /**
             * incumbent = best_known
             */
            $incumbent = $best_known;

            /**
             * start_time = Timer
             * end_time = Timer
             */

            /**
             * Do
             */
            do {
                /**
                 * If iteration Mod 100 = 0 Then
                 *      Application.ScreenUpdating = True
                 *      If best_known.feasible = True Then
                 *           Application.StatusBar = "Reorganizing container " & container_id & ". Best net profit found so far: " & best_known.net_profit & " Distance: " & best_known.total_distance
                 *      Else
                 *           Application.StatusBar = "Reorganizing container " & container_id & ". Best net profit found so far: N/A" & " Distance: " & best_known.total_distance
                 *      End If
                 *      Application.ScreenUpdating = False
                 *
                 *      DoEvents
                 * End If
                 *
                 *
                 * * * Again, this is MS Excel stuff that does stuff on the "frontend".
                 */

                /**
                 * Call PerturbSolution(incumbent, container_id, 0.1 + 0.2 * ((end_time - start_time) / ((solver_options.CPU_time_limit * 0.666) / nonempty_container_cnt)))
                 */
                PerturbSolution($incumbent, $container_id, 0.1 + 0.2 * (($end_time - $start_time) / (($solver_options->CPU_time_limit * 0.666) / $nonempty_container_cnt )), $item_list);

                /**
                 * Call PerturbRotationAndOrderOfItems(incumbent)
                 */
                PerturbRotationAndOrderOfItems($incumbent, $item_list);

                /**
                 * With incumbent
                 */
                $inc = $incumbent;

                /**
                 * For j = 1 To item_list.num_item_types
                 */
                for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                    /**
                     * continue_flag = True
                     */
                    $continue_flag = true;
                    /**
                     * Do While (.unpacked_item_count(.item_type_order(j)) > 0) And (continue_flag = True)
                     *      continue_flag = AddItemToContainer(incumbent, container_id, .item_type_order(j), 1, True)
                     *      'DoEvents
                     * Loop
                     */
                    do {
                        /**
                         * continue_flag = AddItemToContainer(incumbent, container_id, .item_type_order(j), 1, True)
                         */
                        $continue_flag = AddItemToContainer($incumbent, $container_id, $inc->item_type_order[$j], 1, true, $instance, $compatibility_list, $item_list, $solver_options);
                    } while ($inc->unpacked_item_count[$inc->item_type_order[$j]] > 0 && $continue_flag == true);
                }

                /**
                 * .feasible = True
                 */
                $inc->feasible = true;

                /**
                 * For j = 1 To item_list.num_item_types
                 */
                for ($j = 1; $j <= $item_list->num_item_types; ++$j) {
                    /**
                     * If (.unpacked_item_count(j) > 0) And (item_list.item_types(j).mandatory = 1) Then
                     */
                    if ($inc->unpacked_item_count[$j] > 0 && $item_list->item_types[$j]->mandatory == 1) {
                        /**
                         * .feasible = False
                         */
                        $inc->feasible = false;

                        /**
                         * Exit For
                         */
                        break;
                    }
                }

                /**
                 * Call CalculateDistance(incumbent, container_id)
                 */
                CalculateDistance($incumbent, $container_id);

                /**
                 * If ((.feasible = True) And (best_known.feasible = False)) Or _
                 *    ((.feasible = False) And (best_known.feasible = False) And (.total_volume > best_known.total_volume + epsilon)) Or _
                 *    ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit + epsilon)) Or _
                 *    ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_volume < best_known.total_volume - epsilon)) Or _
                 *    ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_volume < best_known.total_volume + epsilon)) And (.total_distance < best_known.total_distance - epsilon) Or _
                 *    ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_volume < best_known.total_volume + epsilon)) And (.total_distance < best_known.total_distance + epsilon) And (.total_x_moment < best_known.total_x_moment - epsilon) Or _
                 *    ((.feasible = True) And (best_known.feasible = True) And (.net_profit > best_known.net_profit - epsilon) And (.total_volume < best_known.total_volume + epsilon)) And (.total_distance < best_known.total_distance + epsilon) And (.total_x_moment < best_known.total_x_moment + epsilon) And (.total_yz_moment < best_known.total_yz_moment - epsilon) Then
                 */

                /*
                 * TODO: triple check this -> really long lines of conditions
                 * */
                if ((($inc->feasible == true) && ($best_known->feasible == false)) ||
                (($inc->feasible == false) && ($best_known->feasible == false) && ($inc->total_volume > $best_known->total_volume + epsilon)) ||
                (($inc->feasible == true) && ($best_known->feasible == true) && ($inc->net_profit > $best_known->net_profit + epsilon)) ||
                (($inc->feasible == true) && ($best_known->feasible == true) && ($inc->net_profit > $best_known->net_profit - epsilon) && ($inc->total_volume < $best_known->total_volume - epsilon)) ||
                (($inc->feasible == true) && ($best_known->feasible == true) && ($inc->net_profit > $best_known->net_profit - epsilon) && ($inc->total_volume < $best_known->total_volume + epsilon)) && ($inc->total_distance < $best_known->total_distance - epsilon) ||
                (($inc->feasible == true) && ($best_known->feasible == true) && ($inc->net_profit > $best_known->net_profit - epsilon) && ($inc->total_volume < $best_known->total_volume + epsilon)) && ($inc->total_distance < $best_known->total_distance + epsilon) && ($inc->total_x_moment < $best_known->total_x_moment - epsilon) ||
                (($inc->feasible == true) && ($best_known->feasible == true) && ($inc->net_profit > $best_known->net_profit - epsilon) && ($inc->total_volume < $best_known->total_volume + epsilon)) && ($inc->total_distance < $best_known->total_distance + epsilon) && ($inc->total_x_moment < $best_known->total_x_moment + epsilon) && ($inc->total_yz_moment < $best_known->total_yz_moment - epsilon))
                {
                    /**
                     * best_known = incumbent
                     */
                    $best_known = $incumbent;

                    /*
                     * ' If best_known.feasible = True Then
                     * '     Application.StatusBar = "Reorganizing container " & container_id & ". Best net profit found so far: " & best_known.net_profit & " Distance: " & best_known.total_distance
                     * ' Else
                     * '     Application.StatusBar = "Reorganizing container " & container_id & ". Best net profit found so far: N/A" & " Distance: " & best_known.total_distance
                     * ' End If
                     * */
                }

                /**
                 * End With
                 */

                /**
                 * iteration = iteration + 1
                 */
                $iteration = $iteration + 1;

                /**
                 * end_time = Timer
                 */

                /**
                 * If end_time < start_time - 0.01 Then
                 *      solver_options.CPU_time_limit = solver_options.CPU_time_limit - (86400 - start_time)
                 *      start_time = end_time
                 * End If
                 *
                 * We'll translate this anyway.
                 */
                if ($end_time < $start_time - 0.01) {
                    $solver_options->CPU_time_limit = $solver_options->CPU_time_limit - (86400 - $start_time);
                    $start_time = $end_time;
                }
            /**
             * Loop While end_time - start_time < (solver_options.CPU_time_limit * 0.666) / nonempty_container_cnt
             */
            } while ($end_time - $start_time < ($solver_options->CPU_time_limit * 0.666) / $nonempty_container_cnt);
        }
    }

    /*
     * 'MsgBox "Iterations performed: " & iteration
     * */
    /**
     * CLP_Solver_Finish:
     *
     * This is defined inside the Sub, so we'll just define it afterwards.
     */
}

/**
 * CLP_Solver_Finish:
 *
 * @param solution_data $best_known
 */
function CLP_Solver_Finish(solution_data $best_known)
{
    /**
     * 'ensure loadability
     */

    /**
     * Dim min_x As Double
     * @var float
     */
    $min_x = 0.0;

    /**
     * Dim min_y As Double
     * @var float
     */
    $min_y = 0.0;

    /**
     * Dim min_z As Double
     * @var float
     */
    $min_z = 0.0;

    /**
     * Dim intersection_right As Double
     * @var float
     */
    $intersection_right = 0.0;

    /**
     * Dim intersection_left As Double
     * @var float
     */
    $intersection_left = 0.0;

    /**
     * Dim intersection_top As Double
     * @var float
     */
    $intersection_top = 0.0;

    /**
     * Dim intersection_bottom As Double
     * @var float
     */
    $intersection_bottom = 0.0;

    /**
     * Dim selected_item_index As Long
     * @var integer
     */
    $selected_item_index = 0;

    /**
     * Dim swap_item As item_in_container
     * @var item_in_container
     */
    $swap_item = new item_in_container();

    /**
     * Dim area_supported As Double
     * @var float
     */
    $area_supported = 0.0;

    /**
     * Dim area_required As Double
     * @var float
     */
    $area_required = 0.0;

    /**
     * Dim support_flag As Boolean
     * @var boolean
     */
    $support_flag = false;

    /**
     * For i = 1 To best_known.num_containers
     */
    for ($i = 1; $i <= $best_known->num_containers; ++$i) {
        /**
         * With best_known.container(i)
         *
         * We'll call this $bkc
         */
        $bkc = $best_known->container[$i];

        /**
         * For j = 1 To .item_cnt
         */
        for ($j = 1; $j <= $bkc->item_cnt; ++$j) {
            /**
             * selected_item_index = 0
             */
            $selected_item_index = 0;

            /**
             * min_x = .width
             */
            $min_x = $bkc->width;

            /**
             * min_y = .height
             */
            $min_y = $bkc->height;

            /**
             * min_z = .length
             */
            $min_z = $bkc->length;

            /**
             * For k = j To .item_cnt
             */
            for ($k = $j; $k <= $bkc->item_cnt; ++$k) {
                /**
                 * If (.items(k).origin_z < min_z - epsilon) Or _
                 *    ((.items(k).origin_z < min_z + epsilon) And (.items(k).origin_y < min_y - epsilon)) Or _
                 *    ((.items(k).origin_z < min_z + epsilon) And (.items(k).origin_y < min_y + epsilon) And (.items(k).origin_x < min_x - epsilon)) Then
                 *
                 */

                /*
                 * TODO: recheck conditions
                 * */
                if (
                    $bkc->items[$k]->origin_z < $min_z - epsilon ||
                    ($bkc->items[$k]->origin_z < $min_z + epsilon && $bkc->items[$k]->origin_y < $min_y - epsilon) ||
                    ($bkc->items[$k]->origin_z < $min_z + epsilon && $bkc->items[$k]->origin_y < $min_y + epsilon && $bkc->items[$k]->origin_x < $min_x - epsilon)
                )
                {
                    /*
                     * 'check for support
                     * */

                    /**
                     * If .items(k).origin_y < epsilon Then
                     */
                    if ($bkc->items[$k]->origin_y < epsilon) {
                        /**
                         * support_flag = True
                         */
                        $support_flag = true;
                    /**
                     * Else
                     */
                    } else {
                        /**
                         * area_supported = 0
                         */
                        $area_supported = 0;

                        /**
                         * area_required = ((.items(k).opposite_x - .items(k).origin_x) * (.items(k).opposite_z - .items(k).origin_z))
                         */
                        $area_required = ($bkc->items[$k]->opposite_x - $bkc->items[$k]->origin_x) * ($bkc->items[$k]->opposite_z - $bkc->items[$k]->origin_z);

                        /**
                         * support_flag = False
                         */
                        $support_flag = false;

                        /**
                         * For l = j - 1 To 1 Step -1
                         */
                        for ($l = $j - 1; $l >= 1; --$j) {
                            /**
                             * If (Abs(.items(k).origin_y - .items(l).opposite_y) < epsilon) Then
                             */
                            if (abs($bkc->items[$k]->origin_y - $bkc->items[$l]->opposite_y) < epsilon) {
                                /*
                                 * 'check for intersection
                                 * */

                                /**
                                 * intersection_right = .items(k).opposite_x
                                 */
                                $intersection_right = $bkc->items[$k]->opposite_x;

                                /**
                                 * If intersection_right > .items(l).opposite_x Then intersection_right = .items(l).opposite_x
                                 */
                                if ($intersection_right > $bkc->items[$l]->opposite_x) {
                                    $intersection_right = $bkc->items[$l]->opposite_x;
                                }

                                /**
                                 * intersection_left = .items(k).origin_x
                                 */
                                $intersection_left = $bkc->items[$k]->origin_x;

                                /**
                                 * If intersection_left > .items(l).origin_x Then intersection_left = .items(l).origin_x
                                 */
                                if ($intersection_left > $bkc->items[$l]->origin_x) {
                                    $intersection_left = $bkc->items[$l]->origin_x;
                                }

                                /**
                                 * intersection_top = .items(k).opposite_z
                                 */
                                $intersection_top = $bkc->items[$k]->opposite_z;

                                /**
                                 * If intersection_top > .items(l).opposite_z Then intersection_top = .items(l).opposite_z
                                 */
                                if ($intersection_top > $bkc->items[$l]->opposite_z) {
                                    $intersection_top = $bkc->items[$l]->opposite_z;
                                }

                                /**
                                 * intersection_bottom = .items(k).origin_z
                                 */
                                $intersection_bottom = $bkc->items[$k]->origin_z;

                                /**
                                 * If intersection_bottom > .items(l).origin_z Then intersection_bottom = .items(l).origin_z
                                 */
                                if ($intersection_bottom > $bkc->items[$l]->origin_z) {
                                    $intersection_bottom = $bkc->items[$l]->origin_z;
                                }

                                /**
                                 * If (intersection_right > intersection_left) And (intersection_top > intersection_bottom) Then
                                 */
                                if ($intersection_right > $intersection_left && $intersection_top > $intersection_bottom) {
                                    /**
                                     * area_supported = area_supported + (intersection_right - intersection_left) * (intersection_top - intersection_bottom)
                                     */
                                    $area_supported = $area_supported + ($intersection_right - $intersection_left) * ($intersection_top - $intersection_bottom);

                                    /**
                                     * If area_supported > area_required - epsilon Then
                                     */
                                    if ($area_supported > $area_required - epsilon) {
                                        /**
                                         * support_flag = True
                                         */
                                        $support_flag = true;

                                        /**
                                         * Exit For
                                         */
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    /**
                     * If support_flag = True Then
                     */
                    if ($support_flag == true) {
                        /**
                         * selected_item_index = k
                         */
                        $selected_item_index = $k;

                        /**
                         * min_x = .items(k).origin_x
                         */
                        $min_x = $bkc->items[$k]->origin_x;

                        /**
                         * min_y = .items(k).origin_y
                         */
                        $min_y = $bkc->items[$k]->origin_y;

                        /**
                         * min_z = .items(k).origin_z
                         */
                        $min_z = $bkc->items[$k]->origin_z;
                    }
                }
            }

            /**
             * If selected_item_index > 0 Then
             */
            if ($selected_item_index > 0) {
                /**
                 * swap_item = .items(selected_item_index)
                 */
                $swap_item = $bkc->items[$selected_item_index];
                $bkc->items[$selected_item_index] = $bkc->items[$j];
                $bkc->items[$j] = $swap_item;
            /**
             * Else
             */
            } else {
                /**
                 * MsgBox ("Loading order could not be constructed.")
                 */
                continue;
            }
        }
        /**
         * End With
         */
    }

    /*
     * 'write the solution
     *
     * 'MsgBox best_known.total_distance
     * */

    /**
     * If best_known.feasible = True Then
     *      reply = MsgBox("CLP Spreadsheet Solver performed " & iteration & " LNS iterations and found a solution with a net profit of " & best_known.net_profit & ". Do you want to overwrite the current solution with the best found solution?", vbYesNo, "CLP Spreadsheet Solver")
     *      If reply = vbYes Then
     *          Call WriteSolution(best_known)
     *      End If
     * ElseIf infeasibility_count > 0 Then
     *      Call WriteSolution(best_known)
     * Else
     *      reply = MsgBox("The best found solution after " & iteration & " LNS iterations does not satisfy all constraints. Do you want to overwrite the current solution with the best found solution?", vbYesNo, "CLP Spreadsheet Solver")
     *      If reply = vbYes Then
     *          Call WriteSolution(best_known)
     *      End If
     * End If
     *
     * * * Writes solution to worksheet if user clicks yes on the MsgBox or infeasibilities exist.
     */

    /*
     * 'Erase the data
     * */

    /**
     * Erase item_list.item_types
     * Erase container_list.container_types
     * Erase compatibility_list.item_to_item
     * Erase compatibility_list.container_to_item
     *
     * For i = 1 To incumbent.num_containers
     *      Erase incumbent.container(i).items
     * Next i
     *
     * Erase incumbent.container
     * Erase incumbent.unpacked_item_count
     *
     * For i = 1 To best_known.num_containers
     *      Erase best_known.container(i).items
     * Next i
     * Erase best_known.container
     * Erase best_known.unpacked_item_count
     *
     * Application.StatusBar = False
     * Application.ScreenUpdating = True
     * Application.Calculation = xlCalculationAutomatic
     *
     * If CheckWorksheetExistence("4.Visualization") Then
     *      ThisWorkbook.Worksheets("4.Visualization").Activate
     * Else
     *      ThisWorkbook.Worksheets("3.Solution").Activate
     * End If
     * Cells(1, 1).Select
     *
     * * * from what I gather, this part is done to free up some memory after everything.
     */
}

/**
 * Sub FeasibilityCheckData(infeasibility_count As Long, infeasibility_string As String)
 *
 */
function FeasibilityCheckData(int $infeasibility_count, string $infeasibility_string, item_list_data $item_list, container_list_data $container_list, instance_data $instance, )
{
    /**
     * Dim i As Long
     * @var integer
     */
    $i = 0;

    /**
     * Dim j As Long
     * @var integer
     */
    $j = 0;

    /**
     * Dim feasibility_flag As Boolean
     * @var boolean
     */
    $feasibility_flag = true;

    /**
     * infeasibility_count = 0
     */
    $infeasibility_count = 0;

    /**
     * infeasibility_string = vbNullString
     */
    $infeasibility_string = "";

    /**
     * Range(ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 9, 1), ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 8 + (4 * item_list.total_number_of_items), 1)).Clear
     *
     * * * Clear some cells.
     */

    /**
     * Dim volume_capacity_required As Double
     * @var float
     */
    $volume_capacity_required = 0.0;

    /**
     * Dim volume_capacity_available As Double
     * @var float
     */
    $volume_capacity_available = 0.0;

    /**
     * Dim weight_capacity_required As Double
     * @var float
     */
    $weight_capacity_required = 0.0;

    /**
     * Dim weight_capacity_available As Double
     * @var float
     */
    $weight_capacity_available = 0.0;

    /**
     * Dim max_width As Double
     * @var float
     */
    $max_width = 0.0;

    /**
     * Dim max_heigth As Double
     * @var float
     */
    $max_heigth = 0.0;

    /**
     * Dim max_length As Double
     * @var float
     */
    $max_length = 0.0;

    /**
     * volume_capacity_required = 0
     *
     * * * already defined
     */

    /**
     * volume_capacity_available = 0
     *
     * * * already defined
     */

    /**
     * weight_capacity_required = 0
     *
     * * * already defined
     */

    /**
     * weight_capacity_available = 0
     *
     * * * already defined
     */

    /**
     * max_width = 0
     *
     * * * already defined
     */

    /**
     * max_heigth = 0
     *
     * * * already defined
     */

    /**
     * max_heigth = 0
     *
     * * * already defined
     */

    /**
     * max_length = 0
     *
     * * * already defined
     */

    /**
     * With item_list
     *
     * * * We'll just call this $il
     */
    $il = $item_list;

    /**
     * For i = 1 To .num_item_types
     */
    for ($i = 1; $i <= $il->num_item_types; ++$i) {
        /**
         * If .item_types(i).mandatory = 1 Then
         */
        if ($il->item_types[$i]->mandatory == 1) {
            /**
             * volume_capacity_required = volume_capacity_required + (.item_types(i).volume * .item_types(i).number_requested)
             */
            $volume_capacity_required = $volume_capacity_required + ($il->item_types[$i]->volume * $il->item_types[$i]->number_requested);

            /**
             * weight_capacity_required = weight_capacity_required + (.item_types(i).weight * .item_types(i).number_requested)
             */
            $weight_capacity_required = $weight_capacity_required + ($il->item_types[$i]->weight)
        }
    }
    /**
     * End With
     */

    /**
     * With container_list
     *
     * $cl should be fine
     */
    $cl = $container_list;

    /**
     * For i = 1 To .num_container_types
     */
    for ($i = 1; $i <= $cl->num_container_types; ++$i) {
        /**
         * If .container_types(i).mandatory >= 0 Then
         */
        if ($cl->container_types[$i]->mandatory >= 0) {
            /**
             * volume_capacity_available = volume_capacity_available + (.container_types(i).volume_capacity * .container_types(i).number_available)
             */
            $volume_capacity_available = $volume_capacity_available + ($cl->container_types[$i]->volume_capacity * $cl->container_types[$i]->number_available);

            /**
             * weight_capacity_available = weight_capacity_available + (.container_types(i).weight_capacity * .container_types(i).number_available)
             */
            $weight_capacity_available = $weight_capacity_available + ($cl->container_types[$i]->weight_capacity * $cl->container_types[$i]->number_available);

            /**
             * If .container_types(i).width > max_width Then max_width = .container_types(i).width
             */
            if ($cl->container_types[$i]->width > $max_width) {
                $max_width = $cl->container_types[$i]->width;
            }

            /**
             * If .container_types(i).height > max_heigth Then max_heigth = .container_types(i).height
             */
            if ($cl->container_types[$i]->height > $max_heigth) {
                $max_heigth = $cl->container_types[$i]->height;
            }

            /**
             * If .container_types(i).length > max_length Then max_length = .container_types(i).length
             */
            if ($cl->container_types[$i]->length > $max_length) {
                $max_length = $cl->container_types[$i]->length;
            }
        }
    }
    /**
     * End With
     */

    /**
     * If volume_capacity_required > volume_capacity_available + epsilon Then
     */
    if ($volume_capacity_required > $volume_capacity_available + epsilon) {
        /**
         * infeasibility_count = infeasibility_count + 1
         */
        $infeasibility_count = $infeasibility_count + 1;

        /**
         * infeasibility_string = infeasibility_string & "The amount of available volume is not enough to pack the mandatory items." & Chr(13)
         */
        $infeasibility_string = $infeasibility_string .  "The amount of available volume is not enough to pack the mandatory items.";

        /**
         * ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 8 + infeasibility_count, 1).Value = "The amount of available volume is not enough to pack the mandatory items."
         *
         * * * Excel stuff
         */
    }

    /**
     * If weight_capacity_required > weight_capacity_available + epsilon Then
     */
    if ($weight_capacity_required > $weight_capacity_available + epsilon) {
        /**
         * infeasibility_count = infeasibility_count + 1
         */
        $infeasibility_count = $infeasibility_count + 1;

        /**
         * infeasibility_string = infeasibility_string & "The amount of available weight capacity is not enough to pack the mandatory items." & Chr(13)
         */
        $infeasibility_string = $infeasibility_string . "The amount of available weight capacity is not enough to pack the mandatory items.";

        /**
         * ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 8 + infeasibility_count, 1).Value = "The amount of available weight capacity is not enough to pack the mandatory items."
         *
         * * * Excel stuff.
         */
    }

    /**
     * With item_list
     * $il
     */
    $il = $item_list;

    /**
     * For i = 1 To .num_item_types
     */
    for ($i = 1; $i <= $il->num_item_types; ++$i) {
        /**
         * With .item_types(i)
         *
         * This will be $ilit
         */
        $ilit = $il->item_types[$i];

        /**
         * If (.mandatory = 1) And (.xy_rotatable = False) And (.yz_rotatable = False) And ((.width > max_width + epsilon) Or (.height > max_heigth + epsilon) Or (.length > max_length + epsilon)) Then
         */
        if (
            $ilit->mandatory == 1 &&
            $ilit->xy_rotatable == false &&
            $ilit->yz_rotatable == false &&
            (
                $ilit->width > $max_width + epsilon ||
                $ilit->height > $max_heigth + epsilon ||
                $ilit->length > $max_length + epsilon
            )
        ) {
            /**
             * infeasibility_count = infeasibility_count + 1
             */
            $infeasibility_count = $infeasibility_count + 1;

            /**
             * If infeasibility_count < 5 Then
             */
            if ($infeasibility_count < 5) {
                /**
                 * infeasibility_string = infeasibility_string & "Item type " & i & " is too large to fit into any container." & Chr(13)
                 */
                $infeasibility_string = $infeasibility_string . "Item type " . $i . " is too large to fit into any container.";
            }

            /**
             * If infeasibility_count = 5 Then
             */
            if ($infeasibility_count == 5) {
                /**
                 * infeasibility_string = infeasibility_string & "More can be found in the list of detected infeasibilities in the solution worksheet." & Chr(13)
                 */
                $infeasibility_string = $infeasibility_string . "More can be found in the list of detected infeasibilities in the solution worksheet.";
            }

            /**
             * ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 8 + infeasibility_count, 1).Value = "Item type " & i & " is too large to fit into any container."
             *
             * * * Excel stuff
             */
        }

        /**
         * If (.mandatory = 1) And (.width > max_width + epsilon) And (.width > max_heigth + epsilon) And (.width > max_length + epsilon) Then
         */
        if (
            $ilit->mandatory == 1 &&
            $ilit->width > $max_width + epsilon &&
            $ilit->width > $max_heigth + epsilon &&
            $ilit->width > $max_length + epsilon
        )
        {
            /**
             * infeasibility_count = infeasibility_count + 1
             */
            $infeasibility_count = $infeasibility_count + 1;

            /**
             * If infeasibility_count < 5 Then
             */
            if ($infeasibility_count < 5) {
                /**
                 * infeasibility_string = infeasibility_string & "Item type " & i & " is too wide to fit into any container." & Chr(13)
                 */
                $infeasibility_string = $infeasibility_string . "Item type " . $i . " is too wide to fit into any container.";
            }

            /**
             * If infeasibility_count = 5 Then
             */
            if ($infeasibility_count == 5) {
                /**
                 * infeasibility_string = infeasibility_string & "More can be found in the list of detected infeasibilities in the solution worksheet." & Chr(13)
                 */
                $infeasibility_string = $infeasibility_string . "More can be found in the list of detected infeasibilities in the solution worksheet.";
            }

            /**
             * ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 8 + infeasibility_count, 1).Value = "Item type " & i & " is too wide to fit into any container."
             */
        }


        /**
         * If (.mandatory = 1) And (.height > max_width + epsilon) And (.height > max_heigth + epsilon) And (.height > max_length + epsilon) Then
         */
        if (
            $ilit->mandatory == 1 &&
            $ilit->height > $max_width + epsilon &&
            $ilit->height > $max_heigth + epsilon &&
            $ilit->height > $max_length + epsilon
        )
        {
            /**
             * infeasibility_count = infeasibility_count + 1
             */
            $infeasibility_count = $infeasibility_count + 1;

            /**
             * If infeasibility_count < 5 Then
             */
            if ($infeasibility_count < 5) {
                /**
                 * infeasibility_string = infeasibility_string & "Item type " & i & " is too tall to fit into any container." & Chr(13)
                 */
                $infeasibility_string = $infeasibility_string . "Item type " . $i . " is too tall to fit into any container.";
            }

            /**
             * If infeasibility_count = 5 Then
             */
            if ($infeasibility_count == 5) {
                /**
                 * infeasibility_string = infeasibility_string & "More can be found in the list of detected infeasibilities in the solution worksheet." & Chr(13)
                 */
                $infeasibility_string = $infeasibility_string . "More can be found in the list of detected infeasibilities in the solution worksheet.";
            }

            /**
             * ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 8 + infeasibility_count, 1).Value = "Item type " & i & " is too tall to fit into any container."
             */
        }

        /**
         * If (.mandatory = 1) And (.length > max_width + epsilon) And (.length > max_heigth + epsilon) And (.length > max_length + epsilon) Then
         */
        if (
            $ilit->mandatory == 1 &&
            $ilit->length > $max_width + epsilon &&
            $ilit->length > $max_heigth + epsilon &&
            $ilit->length > $max_length + epsilon
        )
        {
            /**
             * infeasibility_count = infeasibility_count + 1
             */
            $infeasibility_count = $infeasibility_count + 1;

            /**
             * If infeasibility_count < 5 Then
             */
            if ($infeasibility_count < 5) {
                /**
                 * infeasibility_string = infeasibility_string & "Item type " & i & " is too long to fit into any container." & Chr(13)
                 */
                $infeasibility_string = $infeasibility_string . "Item type " . $i . " is too long to fit into any container.";
            }

            /**
             * If infeasibility_count = 5 Then
             */
            if ($infeasibility_count == 5) {
                /**
                 * infeasibility_string = infeasibility_string & "More can be found in the list of detected infeasibilities in the solution worksheet." & Chr(13)
                 */
                $infeasibility_string = $infeasibility_string . "More can be found in the list of detected infeasibilities in the solution worksheet.";
            }

            /**
             * ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 8 + infeasibility_count, 1).Value = "Item type " & i & " is too long to fit into any container."
             */
        }
        /**
         * End With
         */
    }
    /**
     * End With
     */

    /**
     * If instance.container_item_compatibility_worksheet = True Then
     */
    if ($instance->container_item_compatibility_worksheet == true) {
        /**
         * For i = 1 To item_list.num_item_types
         */
        for ($i = 1; $i <= $item_list->num_item_types; ++$i) {
            /**
             * feasibility_flag = False
             */
            $feasibility_flag = false;

            /**
             * For j = 1 To container_list.num_container_types
             */
            for ($j = 1; $j <= $container_list->num_container_types; ++$j) {
                /**
                 * If compatibility_list.container_to_item(j, i) = True Then
                 */
                if ($compatibility_list->container_to_item[$j][$i] == true) {
                    /**
                     * feasibility_flag = True
                     */
                    $feasibility_flag = true;
                    /**
                     * Exit For
                     */
                    break;
                }
            }

            /**
             * If feasibility_flag = False Then
             */
            if ($feasibility_flag == false) {
                /**
                 * infeasibility_count = infeasibility_count + 1
                 */
                $infeasibility_count = $infeasibility_count + 1;

                /**
                 * If infeasibility_count < 5 Then
                 */
                if ($infeasibility_count < 5) {
                    /**
                     * infeasibility_string = infeasibility_string & "Item type " & i & " is not compatible with any container." & Chr(13)
                     */
                    $infeasibility_string = $infeasibility_string . "Item type " . $i . " is not compatible with any container.";
                }

                /**
                 * If infeasibility_count = 5 Then
                 */
                if ($infeasibility_count == 5) {
                    /**
                     * infeasibility_string = infeasibility_string & "More can be found in the list of detected infeasibilities in the solution worksheet." & Chr(13)
                     */
                    $infeasibility_string = $infeasibility_string . "More can be found in the list of detected infeasibilities in the solution worksheet.";
                }
                /**
                 * ThisWorkbook.Worksheets("3.Solution").Cells(item_list.total_number_of_items + 8 + infeasibility_count, 1).Value = "Item type " & i & " is not compatible with any container."
                 */
            }
        }
    }
}
