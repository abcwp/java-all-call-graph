-- 生成向上的方法完整调用链(指定类名)
select DISTINCT a.full_name,a.simple_name from class_name_impl_dev a join method_call_impl_dev b on a.full_name = b.callee_full_class_name;