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
 * @param solution_data $solution
 * @param int $container_index
 * @param int $item_type_index
 * @param int $add_type
 * @param bool $item_cohesion
 * @param instance_data $instance
 * @param compatibility_data $compatibility_list
 * @param item_list_data $item_list
 */
function AddItemToContainer(solution_data $solution, int $container_index, int $item_type_index, int $add_type, bool $item_cohesion, instance_data $instance, compatibility_data $compatibility_list, item_list_data $item_list)
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
            AddItemToContainer_Finish();
        }
    }
}